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
' 2.11.0
'   2017-09-26 (joe h)    - PO Creation/Modify: Returns valid activity IDs when provided activity ID is invalid (adds to error message). Now fixed.
'                         - Added PeopleSoft_Page_ModalWindow_ExtractSearchTableContents: Generic utility function to read search tables in PS prompts (e.g., valid values for specific fields)
'   2017-09-22 (joe h)    - Major overhaul to Receipt_Create (formerly PurchaseOrder_ProcessReceipt). Page extraction code moved to their own functions
'                              - Receipt Lines now matched by line/schedule. (No more errors when they are out-of-order)
'                              - After receipt, has these been checked for accurancy is valid popup is ignored.
'                              - If valid receipt ID is generated, automatically acknowledges each popup regardless of message
'                              - Checks for duplicate PO Line/schedules before running (breaks the code)
'                              - Continues receiving on items even if one line is not receivable. The error is reported in the Receipt Item error
'                          - PO Creation Updates (CutPO and CreateFromQuote)
'                              - Quote attachment feature added
'                              - Excess Available is acknowledged and ignored
'                              - Create from eQuote PS issues fixed
'                              - Clicks through warning when due date is selected: Due Date selected is a weekend or a public holiday
'                              - Save with Budget Check: Increased timeout waiting period
'                              - Ignores popup if line amount is $0
'                          - Added PeopleSoft_Page_CheckForPopup. Supercedes SuppressPopup(): Checks for buttons, auto-acknowledges, and checks for expected text
'    2017-09-20 (joe h)    - PurchaseOrder_ProcessReceipt: Fixed bug: Subscript out of range error (Caused by previous receipt PO having line item count > current PO)
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
'   - PO Create: Fill xPress Bid Field ID field
'   - PO Create eQuote: Ignores spaces before/after eQuote # match checked
'   - Convert all Qty data types to Currency
'

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
    PopupText As String
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
    PO_HDR_APPROVER_ID As Long
    
    PO_HDR_PO_REF As String  ' NOTE: MAX LENGTH: 30 CHARS
    PO_HDR_COMMENTS As String
    PO_HDR_QUOTE As String
    
    Quote_Attachment_FilePath As String
    
    ' Field Validation Results
    PO_BUSINESS_UNIT_Result As PeopleSoft_Field_ValidationResult
    VENDOR_NAME_SHORT_Result As PeopleSoft_Field_ValidationResult
    PO_HDR_VENDOR_ID_Result As PeopleSoft_Field_ValidationResult
    PO_HDR_VENDOR_LOCATION_Result As PeopleSoft_Field_ValidationResult
    PO_HDR_BUYER_ID_Result As PeopleSoft_Field_ValidationResult
    PO_HDR_APPROVER_ID_Result As PeopleSoft_Field_ValidationResult
    
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
    QTY As Variant ' Decimal Data Type - We must use a fixed point data type here
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

    SCH_DUE_DATE As Date
    SCH_SHIPTO_ID As Long
    
    DIST_BUSINESS_UNIT_PC As String
    DIST_PROJECT_CODE As String
    DIST_ACTIVITY_ID As String
    DIST_LOCATION_ID As Long
    
    SCH_DUE_DATE_Result As PeopleSoft_Field_ValidationResult
    SCH_SHIPTO_ID_Result As PeopleSoft_Field_ValidationResult
    
    DIST_BUSINESS_UNIT_PC_Result As PeopleSoft_Field_ValidationResult
    DIST_PROJECT_CODE_Result As PeopleSoft_Field_ValidationResult
    DIST_ACTIVITY_ID_Result As PeopleSoft_Field_ValidationResult
    DIST_LOCATION_ID_Result As PeopleSoft_Field_ValidationResult
End Type


Type PeopleSoft_PurchaseOrder_Distribution_Fields
    ' PSoft Fields
    BUSINESS_UNIT_PC As String
    PROJECT_CODE As String
    ACTIVITY_ID As String
    LOCATION_ID As Long
    
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

Type PeopleSoft_PurchaseOrder_BudgetCheckResult
    BudgetCheck_HasErrors As Boolean
    BudgetCheck_Errors As PeopleSoft_PurchaseOrder_BudgetCheckErrors
    
    PO_ID As String
    
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
    
    BudgetCheck_Result As PeopleSoft_PurchaseOrder_BudgetCheckResult
    
End Type


' ------------------------------------------------
' PO Create From Quote Fields
' ------------------------------------------------
Type PeopleSoft_PurchaseOrder_CreateFromQuote_LineModification
    PO_Line As Integer
    'PO_Schedule as integer
    
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
    
    BudgetCheck_Result As PeopleSoft_PurchaseOrder_BudgetCheckResult

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
    
    ' PO Defaults
    PO_DUE_DATE As Date
    PO_PROJECT_CODE As String
    
    PO_DUE_DATE_Result As PeopleSoft_Field_ValidationResult
    PO_PROJECT_CODE_Result As PeopleSoft_Field_ValidationResult
    
    ' PO Fields
    PO_HDR_BUYER_ID As Long
    PO_HDR_BUYER_ID_Result As PeopleSoft_Field_ValidationResult
    
    PO_HDR_PO_REF As String
    
    PO_HDR_FLG_SEND_TO_VENDOR As PeopleSoft_Page_CheckboxAction
    
    PO_ChangeOrder_Items() As PeopleSoft_PurchaseOrder_ChangeOrder_Item
    PO_ChangeOrder_ItemCount As Integer
    
    
    ChangeReason As String
    
    
    BudgetCheck_Result As PeopleSoft_PurchaseOrder_BudgetCheckResult
    
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
    
    Receive_Qty As Variant ' Decimal type
    Accept_Qty As Variant ' Decimal type
    
    IsNotReceivable As Boolean ' Returns True if not receivable (receive checkbox is greyed out)
    HasError As Boolean
    ItemError As String
End Type

Type PeopleSoft_Receipt
    PO_BU As String
    PO_ID As String
    
    PO_BU_Result As PeopleSoft_Field_ValidationResult
    
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
    PO_Qty As Double
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
    
    
    Receipt_Qty As Variant ' Decimal type
    Accept_Qty As Variant ' Decimal type
    Receipt_Price As Variant ' Decimal type
    
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
    
    BudgetCheck_Result As PeopleSoft_PurchaseOrder_BudgetCheckResult
    
    HasGlobalError As Boolean
    GlobalError As String
End Type
' ------------------------------------------------
' Constants
' ------------------------------------------------
Private Const URI_BASE As String = "https://erpprd-fnprd.erp.vzwcorp.com/"
'Private Const PS_URI_LOGIN As String = "https://erpprd-fnprd.erp.vzwcorp.com/psc/ps/EMPLOYEE/ERP/c/MANAGE_PURCHASE_ORDERS.PURCHASE_ORDER_EXP.GBL" ' We can use PS page
Private Const PS_URI_LOGIN As String = "https://websso.vzwcorp.com/siteminderagent/forms/vzwsso/sso_login.asp?TARGET=https://websso.vzwcorp.com/profileChk/chkProfile.asp?Orig_Trgt=HTTPS%3a%2f%2ferpprd-fnprd%2eerp%2evzwcorp%2ecom%2fpsp%2fps%2fEMPLOYEE%2fERP%2fh%2f%3ftab%3dDEFAULT"
Private Const PS_URI_PO_EXPRESS As String = "https://erpprd-fnprd.erp.vzwcorp.com/psc/ps/EMPLOYEE/ERP/c/MANAGE_PURCHASE_ORDERS.PURCHASE_ORDER_EXP.GBL"
Private Const PS_URI_RECEIPT_ADD As String = "https://erpprd-fnprd.erp.vzwcorp.com/psc/ps/EMPLOYEE/ERP/c/MANAGE_SHIPMENTS.RECV_PO.GBL"

Private Const TIMEOUT_LONG = 60 * 3 ' 3min
Private Const TIMEOUT_IMPLICIT_WAIT = 3000 '3s

' ------------------------------------------------
' Debug Stuff
' ------------------------------------------------
Private Type PeopleSoft_Debug_Options
    CaptureExceptions As Boolean
    AddFunctionNameToExceptions As Boolean
    TakeScreenshotsAfterException As Boolean
End Type

Private DEBUG_OPTIONS As PeopleSoft_Debug_Options


Public Function GetSeleniumVersion() As String

    Dim assy As New SeleniumWrapper.Assembly
    
    GetSeleniumVersion = assy.GetVersion
    

End Function
Private Sub PeopleSoft_SaveDebugInfo(driver As SeleniumWrapper.WebDriver, Optional prefix As String)

    Dim fileNamePrefix As String, fileHandle As Long
    
    fileNamePrefix = ThisWorkbook.Path & "\" & IIf(prefix <> "", prefix & "_", "") & "PS_" & Format$(Now(), "YYYYmmddHhmmSs")
    driver.captureEntirePageScreenshot fileNamePrefix & "_SS.png"

    fileHandle = FreeFile
    
    Open fileNamePrefix & "_src.html" For Output As #fileHandle
        Write #fileHandle, driver.PageSource
    Close #fileHandle

End Sub


' -----------------------------------------------------------------------------------------------
Public Function PeopleSoft_NewSession(user As String, pass As String) As PeopleSoft_Session

    Dim session As PeopleSoft_Session
    Dim driver As New SeleniumWrapper.WebDriver
    
    ' Setup global debug options
    DEBUG_OPTIONS.CaptureExceptions = True
    DEBUG_OPTIONS.AddFunctionNameToExceptions = False
    
    
    Set session.driver = driver
    
    
    
    session.user = user
    session.pass = pass
    session.loggedIn = False
    
    PeopleSoft_NewSession = session

End Function


Public Function PeopleSoft_Login(ByRef session As PeopleSoft_Session) As Boolean
    
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
            errMsg = driver.findElementByXPath(".//*[starts-with(@id,'ErrorBox')]//p/b").Text
                    
            session.LogonError = "PeopleSoft Login Failed: " & errMsg
        End If
    
        
        session.loggedIn = True
    End If
    
    PeopleSoft_Login = session.loggedIn
    Exit Function
  
ExceptionThrown:
    session.LogonError = "Exception: " & Err.Description
    
    PeopleSoft_Login = False

End Function

Public Function PeopleSoft_NavigateTo_AddPO(ByRef session As PeopleSoft_Session, PO_BU As String, ByRef PO_BU_Result As PeopleSoft_Field_ValidationResult) As Boolean

    
    If DEBUG_OPTIONS.AddFunctionNameToExceptions Then On Error GoTo ExceptionThrown


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
Public Sub PeopleSoft_NavigateTo_ExistingPO(ByRef session As PeopleSoft_Session, PO_BU As String, PO_ID As String)
    
    If DEBUG_OPTIONS.AddFunctionNameToExceptions Then On Error GoTo ExceptionThrown
    
    
    Dim By As New By, Assert As New Assert, Verify As New Verify
    Dim driver As New SeleniumWrapper.WebDriver
    

    
    Set driver = session.driver
    
    
    driver.get PS_URI_PO_EXPRESS
    
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
    
    'PeopleSoft_Page_TypeCalculatedField(driver, fieldElement, fieldValue)
    driver.findElementById("PO_SRCH_OPRID_ENTERED_BY").Clear
    
    
    
    driver.findElementById("#ICSearch").Click
    'driver.runScript "javascript:submitAction_win0(document.win0, '#ICSearch');"
    
    
    PeopleSoft_Page_WaitForProcessing driver, 120
    
    
    Exit Sub
    
    
ExceptionThrown:
    Err.Raise Err.Number, Err.Source, "PeopleSoft_NavigateTo_ExistingPO Exception: " & Err.Description, Err.Helpfile, Err.HelpContext
    
    
End Sub
Public Function PeopleSoft_PurchaseOrder_CutPO(ByRef session As PeopleSoft_Session, ByRef purchaseOrder As PeopleSoft_PurchaseOrder) As Boolean

    If DEBUG_OPTIONS.CaptureExceptions Then On Error GoTo ExceptionThrown


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
    
    
    
    PeopleSoft_Page_SetValidatedField driver, ("VENDOR_VENDOR_NAME_SHORT"), _
        purchaseOrder.PO_Fields.VENDOR_NAME_SHORT, purchaseOrder.PO_Fields.VENDOR_NAME_SHORT_Result
        
    If purchaseOrder.PO_Fields.VENDOR_NAME_SHORT_Result.ValidationFailed Then GoTo ValidationFail
    
    'TODO: PeopleSoft_Page_TypeCalculatedField driver, driver.findElementById("PO_HDR_VENDOR_ID$73$"), "0000000435"
    
    
    Dim vendorLocationText As String
    
    vendorLocationText = driver.findElementById("Z_VNDR_PNLS_WRK_VNDR_LOC").getAttribute("value")
    
    If vendorLocationText = "" Then
    
        If Len(purchaseOrder.PO_Fields.PO_HDR_VENDOR_LOCATION) > 0 Then
            PeopleSoft_Page_SetValidatedField driver, ("Z_VNDR_PNLS_WRK_VNDR_LOC"), _
                purchaseOrder.PO_Fields.PO_HDR_VENDOR_LOCATION, purchaseOrder.PO_Fields.PO_HDR_VENDOR_LOCATION_Result
                
        Else
            purchaseOrder.PO_Fields.PO_HDR_VENDOR_LOCATION_Result.ValidationFailed = True
            purchaseOrder.PO_Fields.PO_HDR_VENDOR_LOCATION_Result.ValidationErrorText = "Vendor location is required"
        End If
        
        
        If purchaseOrder.PO_Fields.PO_HDR_VENDOR_LOCATION_Result.ValidationFailed Then GoTo ValidationFail
    End If
    
    
    PeopleSoft_Page_SetValidatedField driver, ("PO_HDR_BUYER_ID"), _
        CStr(purchaseOrder.PO_Fields.PO_HDR_BUYER_ID), purchaseOrder.PO_Fields.PO_HDR_BUYER_ID_Result
        
    If purchaseOrder.PO_Fields.PO_HDR_BUYER_ID_Result.ValidationFailed Then GoTo ValidationFail
    
    
    If Len(purchaseOrder.PO_Fields.PO_HDR_PO_REF) > 0 Then
        driver.findElementById("PO_HDR_PO_REF").Clear
        driver.findElementById("PO_HDR_PO_REF").SendKeys purchaseOrder.PO_Fields.PO_HDR_PO_REF
    End If
    
    
    
    ' -------------------------------------------------------------------
    ' Begin - Header Section
    ' -------------------------------------------------------------------
    If Len(purchaseOrder.PO_Fields.PO_HDR_QUOTE) > 0 Then
        ' Only if quote field provided
    
        driver.findElementById("PO_PNLS_WRK_GOTO_HDR_DTL").Click
        'javascript:hAction_win0(document.win0,'PO_PNLS_WRK_GOTO_HDR_DTL', 0, 0, 'Header Details', false, true);
        
        'driver.waitForPageToLoad 5000
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
    Dim fillResult As Boolean
    fillResult = PeopleSoft_PurchaseOrder_PO_Fill_Comments_Page(driver, purchaseOrder.PO_Fields)
    If Not fillResult Then GoTo ValidationFail ' TODO: Add .HasValidationError calculation
    
    ' -------------------------------------------------------------------
    ' Begin - Comments Section
    ' -------------------------------------------------------------------
    'If Len(purchaseOrder.PO_Fields.PO_HDR_COMMENTS) > 0 Then ' Only go into comments section if commend is given
    '    driver.findElementById("COMM_WRK1_COMMENTS_PB").Click
    '    'driver.runScript "javascript:hAction_win0(document.win0,'COMM_WRK1_COMMENTS1_PB', 0, 0, 'Edit Comments', false, true);"
    '
    '    'driver.waitForPageToLoad 5000
    '    PeopleSoft_Page_WaitForProcessing driver
    '
    '    If False Then ' No more suggested approver
    '        driver.waitForElementPresent "css=#PO_HDR_Z_SUG_APPRVR"
    '
    '
    '        PeopleSoft_Page_SetValidatedField driver, ("PO_HDR_Z_SUG_APPRVR"), _
    '            CStr(purchaseOrder.PO_Fields.PO_HDR_APPROVER_ID), purchaseOrder.PO_Fields.PO_HDR_APPROVER_ID_Result
    '
    '        If purchaseOrder.PO_Fields.PO_HDR_APPROVER_ID_Result.ValidationFailed Then GoTo ValidationFail
    '    End If
    '
    '    If Len(purchaseOrder.PO_Fields.PO_HDR_COMMENTS) > 0 Then
    '        driver.findElementById("COMM_WRK1_COMMENTS_2000$0").Clear
    '        driver.findElementById("COMM_WRK1_COMMENTS_2000$0").SendKeys purchaseOrder.PO_Fields.PO_HDR_COMMENTS
    '    End If
    '
    '    PeopleSoft_Page_WaitForProcessing driver
    '
   '
   '     driver.findElementById("#ICSave").Click
   '     'driver.runScript "javascript:submitAction_win0(document.win0, '#ICSave');" ' work-around - Clicks 'Save'
   '
   '     PeopleSoft_Page_WaitForProcessing driver
   '
   ' End If
    ' -------------------------------------------------------------------
    ' End - Comments Section
    ' -------------------------------------------------------------------
    
    Dim PO_Line As Integer
    Dim PO_LineCount As Integer
    Dim PO_pageLineIndex As Integer
    Dim PO_pageScheduleIndex As Integer
    Dim PO_Line_Schedule As Integer
    
    
    ' Calculate and fill defaults if the PO has more than one line
    If purchaseOrder.PO_LineCount > 1 Then
        purchaseOrder.PO_Defaults = PeopleSoft_PurchaseOrder_PO_Defaults_AutoCalc(purchaseOrder)
        PeopleSoft_PurchaseOrder_PO_Defaults_Fill driver, purchaseOrder.PO_Defaults
        
        ' Begin - Transfer validation errors from defaults to each line/schedule
        If purchaseOrder.PO_Defaults.HasValidationError Then
            For PO_Line = 1 To purchaseOrder.PO_LineCount
                For PO_Line_Schedule = 1 To purchaseOrder.PO_Lines(PO_Line).ScheduleCount
                
                    With purchaseOrder.PO_Lines(PO_Line).Schedules(PO_Line_Schedule)
                        .ScheduleFields.DUE_DATE_Result = purchaseOrder.PO_Defaults.SCH_DUE_DATE_Result
                        .ScheduleFields.SHIPTO_ID_Result = purchaseOrder.PO_Defaults.SCH_SHIPTO_ID_Result
                        
                        .DistributionFields.BUSINESS_UNIT_PC_Result = purchaseOrder.PO_Defaults.DIST_BUSINESS_UNIT_PC_Result
                        .DistributionFields.PROJECT_CODE_Result = purchaseOrder.PO_Defaults.DIST_PROJECT_CODE_Result
                        .DistributionFields.ACTIVITY_ID_Result = purchaseOrder.PO_Defaults.DIST_ACTIVITY_ID_Result
                        .DistributionFields.LOCATION_ID_Result = purchaseOrder.PO_Defaults.DIST_LOCATION_ID_Result
                    End With
                    
                    With purchaseOrder.PO_Defaults
                        purchaseOrder.PO_Lines(PO_Line).HasValidationError = purchaseOrder.PO_Lines(PO_Line).HasValidationError _
                            Or .SCH_DUE_DATE_Result.ValidationFailed _
                            Or .SCH_SHIPTO_ID_Result.ValidationFailed _
                            Or .DIST_BUSINESS_UNIT_PC_Result.ValidationFailed _
                            Or .DIST_PROJECT_CODE_Result.ValidationFailed _
                            Or .DIST_ACTIVITY_ID_Result.ValidationFailed _
                            Or .DIST_LOCATION_ID_Result.ValidationFailed
                    End With
                    
                Next PO_Line_Schedule
            Next PO_Line
            
             GoTo ValidationFail
        End If
        ' End - Transfer validation errors from defaults to each line/schedule
    
        
    End If
    
    ' -------------------------------------------------------------------
    ' Begin - Add individual lines to PO
    ' -------------------------------------------------------------------
    
    
    
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
    
    
    
    'Dim anyLineHasValidationError As Boolean
    
    'anyLineHasValidationError = False
    
    For PO_Line = 1 To PO_LineCount
        'Debug.Print "Line: " & PO_line
        
        PeopleSoft_PurchaseOrder_Fill_PO_Line driver, purchaseOrder, PO_Line, PO_pageScheduleIndex
        If purchaseOrder.HasError Then GoTo ValidationFail
        
        'If purchaseOrder.PO_Lines(PO_Line).HasValidationError Then anyLineHasValidationError = True
        If purchaseOrder.PO_Lines(PO_Line).HasValidationError Then GoTo ValidationFail
        
        
        PO_pageScheduleIndex = PO_pageScheduleIndex + purchaseOrder.PO_Lines(PO_Line).ScheduleCount
    Next PO_Line
    
    
    'If anyLineHasValidationError Then GoTo ValidationFail
    
       
    driver.runScript "javascript:submitAction_win0(document.win0,'CALCULATE_TAXES');" ' Fix for 2.9.1.1  due to PS upgrade
    'driver.findElementById("CALCULATE_TAXES").Click
    'driver.runScript "javascript:hAction_win0(document.win0,'CALCULATE_TAXES', 0, 0, 'Calculate', false, true);"
    
    PeopleSoft_Page_WaitForProcessing driver

    
    Dim amntStr As String
    
    ' Total
    amntStr = driver.findElementById("PO_PNLS_WRK_PO_AMT_TTL").Text
    purchaseOrder.PO_AMNT_TOTAL = CurrencyFromString(amntStr)
    
    ' Total w/o Taxes, Freight and Misc
    amntStr = driver.findElementById("PO_PNLS_WRK_MERCH_AMT_TTL").Text
    purchaseOrder.PO_AMNT_MERCH_TOTAL = CurrencyFromString(amntStr)
    
    ' Taxes, Freight and Misc
    amntStr = driver.findElementById("PO_PNLS_WRK_ADJ_AMT_TTL_LBL").Text
    purchaseOrder.PO_AMNT_FTM_TOTAL = CurrencyFromString(amntStr)
    
    
    
    Dim result As Boolean
    
    result = PeopleSoft_PurchaseOrder_SaveWithBudgetCheck(driver, purchaseOrder.BudgetCheck_Result)
    
    If result = False Then
        purchaseOrder.GlobalError = purchaseOrder.BudgetCheck_Result.GlobalError
        purchaseOrder.HasError = purchaseOrder.BudgetCheck_Result.HasGlobalError
        
        PeopleSoft_PurchaseOrder_CutPO = False
        Exit Function
    End If
    
    purchaseOrder.PO_ID = purchaseOrder.BudgetCheck_Result.PO_ID
    
    PeopleSoft_PurchaseOrder_CutPO = True
    Exit Function
    
    
ValidationFail:
    PeopleSoft_PurchaseOrder_CutPO = False
    Exit Function
    
ExceptionThrown:
    purchaseOrder.HasError = True
    purchaseOrder.GlobalError = "Exception: " & Err.Description
    
    PeopleSoft_PurchaseOrder_CutPO = False


End Function
Public Function PeopleSoft_PurchaseOrder_CreateFromQuote(ByRef session As PeopleSoft_Session, ByRef poCFQ As PeopleSoft_PurchaseOrder_CreateFromQuoteParams) As Boolean

    If DEBUG_OPTIONS.CaptureExceptions Then On Error GoTo ExceptionThrown
    
    
    Debug.Print "PeopleSoft_PurchaseOrder_CreateFromQuote"
    


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
    
    ' Test fail point
    driver.findElementById "this element does not exist"
    
    
    PeopleSoft_Page_SetValidatedField driver, ("PO_HDR_BUYER_ID"), CStr(poCFQ.PO_Fields.PO_HDR_BUYER_ID), poCFQ.PO_Fields.PO_HDR_BUYER_ID_Result
    If poCFQ.PO_Fields.PO_HDR_BUYER_ID_Result.ValidationFailed Then GoTo ValidationFail
    
    
    If Len(poCFQ.PO_Fields.PO_HDR_PO_REF) > 0 Then
        driver.findElementById("PO_HDR_PO_REF").Clear
        driver.findElementById("PO_HDR_PO_REF").SendKeys poCFQ.PO_Fields.PO_HDR_PO_REF
    End If
    
    
    'Dim elemSelect As SeleniumWrapper.Select
    Dim elemSelect As SeleniumWrapper.WebElement
    

    ' Select CopyFrom eQuote and force load next page
    Set elemSelect = driver.findElementById("PO_COPY_TMPLT_W_COPY_PO_FROM")
    elemSelect.AsSelect.selectByText "eQuote"
    PeopleSoft_Page_WaitForProcessing driver
    driver.Wait 1000
    driver.runScript "javascript: var elem = document.getElementById('PO_COPY_TMPLT_W_COPY_PO_FROM'); addchg_win0(elem); submitAction_win0(elem.form,elem.name);"
    PeopleSoft_Page_WaitForProcessing driver
    

   
    ' <h1 class="PSSRCHTITLE">Create from Quote</h1>
    driver.waitForElementPresent "xpath=.//*[text()='Create from Quote']"
    

    ' Type Vendor ID
    PeopleSoft_Page_SetValidatedField driver, ("Z_E_QT_WRK_VENDOR_ID"), Format(poCFQ.PO_Fields.PO_HDR_VENDOR_ID, "0000000000"), poCFQ.PO_Fields.PO_HDR_VENDOR_ID_Result
    If poCFQ.PO_Fields.PO_HDR_VENDOR_ID_Result.ValidationFailed Then GoTo ValidationFail
    
    ' Type Quote Number
    PeopleSoft_Page_SetValidatedField driver, ("Z_E_QT_WRK_Z_QUOTE_NBR"), poCFQ.E_QUOTE_NBR, poCFQ.E_QUOTE_NBR_Result
    If poCFQ.E_QUOTE_NBR_Result.ValidationFailed Then GoTo ValidationFail
    
    ' Click Search
    driver.findElementById("Z_E_QT_WRK_REFRESH").Click
    PeopleSoft_Page_WaitForProcessing driver
    
    
    Dim loadedQuoteNbr As String
    loadedQuoteNbr = driver.findElementById("Z_E_QT_CPPO_VW_Z_QUOTE_NBR$0").Text

    ' Sanity check: check if loaded quote is the same as the provided E_QUOTE_NBR
    If loadedQuoteNbr <> poCFQ.E_QUOTE_NBR Then
        poCFQ.HasError = True
        poCFQ.GlobalError = "Sanity check failed: quote # mismatch. Quote # on page '" & loadedQuoteNbr & "' does not match provided quote # '" & poCFQ.E_QUOTE_NBR & "'"
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
    
    If validLineModCount > 0 Then
        ' Begin - Sort Line Modifications by line #
        'Dim PO_LineMod_SortedIdx() As Integer
        'ReDim PO_LineMod_SortedIdx(1 To poCFQ.PO_LineModCount) As Integer
        
        'For i = 1 To poCFQ.PO_LineModCount: PO_LineMod_SortedIdx(i) = i: Next
        
        ' Bubble-sort algorithm
        'For i = 1 To poCFQ.PO_LineModCount
        '    For j = i + 1 To poCFQ.PO_LineModCount
        '        If poCFQ.PO_LineMods(PO_LineMod_SortedIdx(j)).PO_Line < poCFQ.PO_LineMods(PO_LineMod_SortedIdx(i)).PO_Line Then
        '            tmpIdx = PO_LineMod_SortedIdx(i)
        '            PO_LineMod_SortedIdx(i) = PO_LineMod_SortedIdx(j)
        '            PO_LineMod_SortedIdx(j) = tmpIdx
        '        End If
         '   Next j
        'Next i
        ' End - Sort Line Modifications by line #
        
        
        'For i = 1 To poCFQ.PO_LineModCount
        '    PO_LineMod = PO_LineMod_SortedIdx(i)
        '    driver.runScript "javascript:hAction_win0(document.win0,'PO_PNLS_PB_EXPAND_PB$" & (poCFQ.PO_LineMods(PO_LineMod).PO_Line - 1) & "', 0, 0, 'Expand Schedule Section', false, true);"
        '    PeopleSoft_Page_WaitForProcessing driver
        '    driver.runScript "javascript:hAction_win0(document.win0,'PO_PNLS_PB_EXPAND_PB$232$$" & (poCFQ.PO_LineMods(PO_LineMod).PO_Line - 1) & "', 0, 0, 'Expand Distribution Section', false, true);"
        '    PeopleSoft_Page_WaitForProcessing driver
        'Next i
        
        
        ' Expand All
        driver.runScript "javascript:submitAction_win0(document.win0,'PO_PNLS_PB_EXPAND_ALL_PB', 0, 0, 'Expand All', false, true);" ' Fix for 2.9.1.1  due to PS upgrade
        'driver.runScript "javascript:hAction_win0(document.win0,'PO_PNLS_PB_EXPAND_ALL_PB', 0, 0, 'Expand All', false, true);"
        PeopleSoft_Page_WaitForProcessing driver
        
        For PO_LineMod = 1 To poCFQ.PO_LineModCount
            'PO_LineMod = PO_LineMod_SortedIdx(i)
            
            
            If poCFQ.PO_LineMods(PO_LineMod).PO_Line > 0 Then
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
    amntStr = driver.findElementById("PO_PNLS_WRK_PO_AMT_TTL").Text
    poCFQ.PO_AMNT_TOTAL = CurrencyFromString(amntStr)
    
    ' Total w/o Taxes, Freight and Misc
    amntStr = driver.findElementById("PO_PNLS_WRK_MERCH_AMT_TTL").Text
    poCFQ.PO_AMNT_MERCH_TOTAL = CurrencyFromString(amntStr)
    
    ' Taxes, Freight and Misc
    amntStr = driver.findElementById("PO_PNLS_WRK_ADJ_AMT_TTL_LBL").Text
    poCFQ.PO_AMNT_FTM_TOTAL = CurrencyFromString(amntStr)
    
    'poCFQ.PO_ID = "SUCCESS"
    'PeopleSoft_PurchaseOrder_CreateFromQuote = True
    'Exit Function
    
    Dim result As Boolean
    
    result = PeopleSoft_PurchaseOrder_SaveWithBudgetCheck(driver, poCFQ.BudgetCheck_Result)
    
    If result = False Then
        poCFQ.GlobalError = poCFQ.BudgetCheck_Result.GlobalError
        poCFQ.HasError = poCFQ.BudgetCheck_Result.HasGlobalError
        
        PeopleSoft_PurchaseOrder_CreateFromQuote = False
        Exit Function
    End If
    
    poCFQ.PO_ID = poCFQ.BudgetCheck_Result.PO_ID
    
    PeopleSoft_PurchaseOrder_CreateFromQuote = True
    Exit Function
    
    
ValidationFail:
    PeopleSoft_SaveDebugInfo driver, "eQuote"
    PeopleSoft_PurchaseOrder_CreateFromQuote = False
    Exit Function
    
ExceptionThrown:
    PeopleSoft_SaveDebugInfo driver, "eQuote"
    poCFQ.HasError = True
    poCFQ.GlobalError = "Exception: " & Err.Description
    
    PeopleSoft_PurchaseOrder_CreateFromQuote = False


End Function
Public Function PeopleSoft_PurchaseOrder_PO_Defaults_AutoCalc(purchaseOrder As PeopleSoft_PurchaseOrder) As PeopleSoft_PurchaseOrder_PO_Defaults

    ' Auto calculates PO defaults. A field has a default value when all PO lines/schedules/distributions have the same value

    Dim PO_Defaults As PeopleSoft_PurchaseOrder_PO_Defaults

    Dim PO_Line As Integer, PO_Line_Schedule As Integer
    
    PO_Defaults.SCH_DUE_DATE = 0
    PO_Defaults.SCH_SHIPTO_ID = 0
    PO_Defaults.DIST_BUSINESS_UNIT_PC = ""
    PO_Defaults.DIST_PROJECT_CODE = ""
    PO_Defaults.DIST_ACTIVITY_ID = ""
    PO_Defaults.DIST_LOCATION_ID = 0
    
   
    For PO_Line = 1 To purchaseOrder.PO_LineCount
        For PO_Line_Schedule = 1 To purchaseOrder.PO_Lines(PO_Line).ScheduleCount
        
            If PO_Line = 1 And PO_Line_Schedule = 1 Then
                PO_Defaults.SCH_DUE_DATE = purchaseOrder.PO_Lines(PO_Line).Schedules(PO_Line_Schedule).ScheduleFields.DUE_DATE
                PO_Defaults.SCH_SHIPTO_ID = purchaseOrder.PO_Lines(PO_Line).Schedules(PO_Line_Schedule).ScheduleFields.SHIPTO_ID
                PO_Defaults.DIST_BUSINESS_UNIT_PC = purchaseOrder.PO_Lines(PO_Line).Schedules(PO_Line_Schedule).DistributionFields.BUSINESS_UNIT_PC
                PO_Defaults.DIST_PROJECT_CODE = purchaseOrder.PO_Lines(PO_Line).Schedules(PO_Line_Schedule).DistributionFields.PROJECT_CODE
                PO_Defaults.DIST_ACTIVITY_ID = purchaseOrder.PO_Lines(PO_Line).Schedules(PO_Line_Schedule).DistributionFields.ACTIVITY_ID
                PO_Defaults.DIST_LOCATION_ID = purchaseOrder.PO_Lines(PO_Line).Schedules(PO_Line_Schedule).DistributionFields.LOCATION_ID
            Else
                If PO_Defaults.SCH_DUE_DATE <> purchaseOrder.PO_Lines(PO_Line).Schedules(PO_Line_Schedule).ScheduleFields.DUE_DATE Then
                    PO_Defaults.SCH_DUE_DATE = 0
                End If
                If PO_Defaults.SCH_SHIPTO_ID <> purchaseOrder.PO_Lines(PO_Line).Schedules(PO_Line_Schedule).ScheduleFields.SHIPTO_ID Then
                    PO_Defaults.SCH_SHIPTO_ID = 0
                End If
                If PO_Defaults.DIST_BUSINESS_UNIT_PC <> purchaseOrder.PO_Lines(PO_Line).Schedules(PO_Line_Schedule).DistributionFields.BUSINESS_UNIT_PC Then
                    PO_Defaults.DIST_BUSINESS_UNIT_PC = ""
                End If
                If PO_Defaults.DIST_PROJECT_CODE <> purchaseOrder.PO_Lines(PO_Line).Schedules(PO_Line_Schedule).DistributionFields.PROJECT_CODE Then
                    PO_Defaults.DIST_PROJECT_CODE = ""
                End If
                If PO_Defaults.DIST_ACTIVITY_ID <> purchaseOrder.PO_Lines(PO_Line).Schedules(PO_Line_Schedule).DistributionFields.ACTIVITY_ID Then
                    PO_Defaults.DIST_ACTIVITY_ID = ""
                End If
                If PO_Defaults.DIST_LOCATION_ID <> purchaseOrder.PO_Lines(PO_Line).Schedules(PO_Line_Schedule).DistributionFields.LOCATION_ID Then
                    PO_Defaults.DIST_LOCATION_ID = 0
                End If
            End If
            
        Next PO_Line_Schedule
    Next PO_Line
    
    
    If PO_Defaults.DIST_PROJECT_CODE = "" Then ' Activity & Location default requires as project code default
        PO_Defaults.DIST_ACTIVITY_ID = ""
        PO_Defaults.DIST_LOCATION_ID = 0
    End If
        
    
    PeopleSoft_PurchaseOrder_PO_Defaults_AutoCalc = PO_Defaults

End Function
Private Function PeopleSoft_PurchaseOrder_PO_Defaults_Fill(driver As SeleniumWrapper.WebDriver, PO_Defaults As PeopleSoft_PurchaseOrder_PO_Defaults) As Boolean

    If DEBUG_OPTIONS.AddFunctionNameToExceptions Then On Error GoTo ExceptionThrown

    Dim isAnyDefaultSpecified As Boolean
    
    Dim PopupText As String
    
    isAnyDefaultSpecified = False
    
    If PO_Defaults.SCH_DUE_DATE > 0 Then isAnyDefaultSpecified = True
    If PO_Defaults.SCH_SHIPTO_ID > 0 Then isAnyDefaultSpecified = True
    If Len(PO_Defaults.DIST_BUSINESS_UNIT_PC) > 0 Then isAnyDefaultSpecified = True
    If Len(PO_Defaults.DIST_PROJECT_CODE) > 0 Then isAnyDefaultSpecified = True
    If Len(PO_Defaults.DIST_ACTIVITY_ID) > 0 Then isAnyDefaultSpecified = True
    If PO_Defaults.DIST_LOCATION_ID > 0 Then isAnyDefaultSpecified = True


    ' Need to implement expense chartfields:
    ' Exp PC BU: Z_EXP_PC_BU$0
    ' Exp Project ID: PROJECT_ID_2$0
    ' Exp Activity ID: ACTIVITY_ID_2$0
    ' Exp: Ship to ID: PO_DFLT_DISTRIB_SHIPTO_ID_DEFAULT$0
    ' Exp: Location: PO_DFLT_DISTRIB_Z_EXP_CF1$0
    ' Exp Cost Center: PO_DFLT_DISTRIB_Z_EXP_DEPTID$0
    ' Exp Product Code: PO_DFLT_DISTRIB_Z_EXP_PROD$0
    
    If isAnyDefaultSpecified Then
     
        driver.findElementById("PO_PNLS_WRK_GOTO_DEFAULTS").Click
        'javascript:hAction_win0(document.win0,'PO_PNLS_WRK_GOTO_DEFAULTS', 0, 0, 'Header Details', false, true);
        
         PeopleSoft_Page_WaitForProcessing driver
         
         
         PopupText = PeopleSoft_Page_SuppressPopup(driver, vbOK)
         
         If Len(PopupText) > 0 Then
            PO_Defaults.GlobalError = PopupText
            PO_Defaults.HasGlobalError = True
         
            PeopleSoft_PurchaseOrder_PO_Defaults_Fill = False
            Exit Function
         End If
        
        'driver.waitForElementPresent "css=#PO_HDR_Z_QUOTE_NBR"
            
        If PO_Defaults.SCH_DUE_DATE > 0 Then
            PeopleSoft_Page_SetValidatedField driver:=driver, fieldElementID:=("PO_DFLT_TBL_DUE_DT"), fieldValue:=Format(PO_Defaults.SCH_DUE_DATE, "mm/dd/yyyy"), _
                validationResult:=PO_Defaults.SCH_DUE_DATE_Result, _
                expectedPopupContents:="*Due Date selected is a weekend or a public holiday*"
            
            If PO_Defaults.SCH_DUE_DATE_Result.ValidationFailed Then GoTo ValidationFail
        End If
    
        ' Expand Distributions
        ' driver.runscript "javascript:submitAction_win0(document.win0,'PO_DFLT_DISTRIB$htab$0');"
        PeopleSoft_Page_WaitForProcessing driver
        

        
        If Len(PO_Defaults.DIST_BUSINESS_UNIT_PC) > 0 Then
            PeopleSoft_Page_SetValidatedField driver, ("BUSINESS_UNIT_PC$0"), PO_Defaults.DIST_BUSINESS_UNIT_PC, PO_Defaults.DIST_BUSINESS_UNIT_PC_Result
            If PO_Defaults.DIST_BUSINESS_UNIT_PC_Result.ValidationFailed Then GoTo ValidationFail
        End If
        
        If Len(PO_Defaults.DIST_PROJECT_CODE) > 0 Then
            PeopleSoft_Page_SetValidatedField driver, ("PROJECT_ID$0"), PO_Defaults.DIST_PROJECT_CODE, PO_Defaults.DIST_PROJECT_CODE_Result
            If PO_Defaults.DIST_PROJECT_CODE_Result.ValidationFailed Then GoTo ValidationFail
        End If
        
        If Len(PO_Defaults.DIST_ACTIVITY_ID) > 0 Then
            PeopleSoft_PurchaseOrder_SetValidatedField_ActivityID driver, _
                "ACTIVITY_ID$0", PO_Defaults.DIST_ACTIVITY_ID, PO_Defaults.DIST_ACTIVITY_ID_Result, "ACTIVITY_ID$prompt$0"
        
            If PO_Defaults.DIST_ACTIVITY_ID_Result.ValidationFailed Then GoTo ValidationFail
            
            'PeopleSoft_Page_SetValidatedField driver, _
            '    ("ACTIVITY_ID$0"), _
            '    PO_Defaults.DIST_ACTIVITY_ID, _
            '    PO_Defaults.DIST_ACTIVITY_ID_Result
           '
           ' If PO_Defaults.DIST_ACTIVITY_ID_Result.ValidationFailed Then
           '     If InStr(1, PO_Defaults.DIST_ACTIVITY_ID_Result.ValidationErrorText, "Invalid value") > 0 Then
           '         Dim tmpFVR As PeopleSoft_Field_ValidationResult
           '         Dim activityListString As String
           '
           '         PeopleSoft_Page_SetValidatedField driver, ("ACTIVITY_ID$0"), "", tmpFVR, False
           '
           '         activityListString = PeopleSoft_PurchaseOrder_Extract_ActivityIDs(driver, "ACTIVITY_ID$prompt$0")
           '
           '         If Len(activityListString) > 0 Then PO_Defaults.DIST_ACTIVITY_ID_Result.ValidationErrorText = "Invalid activity ID. Valid values are as follows: " & activityListString
   
            '    End If
                
            '    GoTo ValidationFail
            'End If
        End If
        
        
        
        
        If PO_Defaults.SCH_SHIPTO_ID > 0 Then
            PeopleSoft_Page_SetValidatedField driver, ("PO_DFLT_DISTRIB_SHIPTO_ID$0"), CStr(PO_Defaults.SCH_SHIPTO_ID), PO_Defaults.SCH_SHIPTO_ID_Result
            If PO_Defaults.SCH_SHIPTO_ID_Result.ValidationFailed Then GoTo ValidationFail
        End If
        
        If PO_Defaults.DIST_LOCATION_ID > 0 Then
            PeopleSoft_Page_SetValidatedField driver, ("LOCATION$0"), CStr(PO_Defaults.DIST_LOCATION_ID), PO_Defaults.DIST_LOCATION_ID_Result
            If PO_Defaults.DIST_LOCATION_ID_Result.ValidationFailed Then GoTo ValidationFail
        End If
        
        
        
        driver.findElementById("#ICSave").Click
        'driver.runScript "javascript:submitAction_win0(document.win0, '#ICSave');" ' work-around - Clicks 'Save'
        
        PeopleSoft_Page_WaitForProcessing driver, TIMEOUT_LONG
        
    End If
    
    
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
    
    
    If DEBUG_OPTIONS.AddFunctionNameToExceptions Then On Error GoTo ExceptionThrown
    
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
            driver.findElementById("COMM_WRK1_COMMENTS_2000$0").Clear
            driver.findElementById("COMM_WRK1_COMMENTS_2000$0").SendKeys poFields.PO_HDR_COMMENTS
            PeopleSoft_Page_WaitForProcessing driver
        End If
        
        
        
        ' If quote file provided -> attach quote
        If Len(poFields.Quote_Attachment_FilePath) > 0 Then
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
                poFields.Quote_Attachment_FilePath_Result.ValidationErrorText = "Attach quote failed: Unexpected popup: " & popupCheckResult.PopupText
                GoTo ValidationFail
            End If
                        
            
            ' Check if file was successfully uploaded
            Dim uploadedFilename As String
            Set we = driver.findElementById("PV_ATTACH_WRK_ATTACHUSERFILE$0")
            uploadedFilename = we.Text
            
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
            pageColName = weCollection.Item(i).Text
            
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
            returnColNames(i) = weCollection.Item(i).Text
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
                ret(i + 1) = we.Text
            Else
                ret(i + 1, j) = we.Text
            End If
        Next j
    Next i
    
    
    driver.selectFrame "relative=top"
    driver.runScript "javascript:doCloseModal(" & modalIndex & ");"
    
    PeopleSoft_Page_ModalWindow_ExtractSearchTableContents = ret


End Function
Public Function PeopleSoft_PurchaseOrder_Fill_PO_Line(driver As SeleniumWrapper.WebDriver, ByRef purchaseOrder As PeopleSoft_PurchaseOrder, PO_Line As Integer, ByVal PO_pageScheduleIndex As Integer) As Boolean

    Debug.Assert PO_Line > 0 And PO_Line <= purchaseOrder.PO_LineCount
    

    
    On Error GoTo ExceptionThrown
    
    
    Dim PO_Line_Schedule As Integer, PO_Line_ScheduleCount As Integer
    
    
    ' Begin - Enter Line Fields
    PeopleSoft_Page_SetValidatedField driver, _
        ("PO_LINE_INV_ITEM_ID$" & (PO_Line - 1)), _
        purchaseOrder.PO_Lines(PO_Line).LineFields.PO_LINE_ITEM_ID, _
        purchaseOrder.PO_Lines(PO_Line).LineFields.PO_LINE_ITEM_ID_Result
    
    If purchaseOrder.PO_Lines(PO_Line).LineFields.PO_LINE_ITEM_ID_Result.ValidationFailed Then GoTo ValidationFail
    
    
    Dim tmpValResult As PeopleSoft_Field_ValidationResult
    
        PeopleSoft_Page_SetValidatedField driver, _
            ("PO_LINE_DESCR254_MIXED$" & (PO_Line - 1)), _
            purchaseOrder.PO_Lines(PO_Line).LineFields.PO_LINE_DESC, _
            tmpValResult
        
        
    If tmpValResult.ValidationFailed Then GoTo ValidationFail


    'PeopleSoft_Page_SetValidatedField  driver, _
    '    driver.findElementById("PO_PNLS_WRK_QTY_PO$" & (PO_Line - 1)), _
    '    CStr(purchaseOrder.PO_Lines(PO_Line).LineFields.PO_LINE_QTY), _
    '    purchaseOrder.PO_Lines(PO_Line).LineFields.PO_LINE_QTY_Result
   '
    'If purchaseOrder.PO_Lines(PO_Line).LineFields.PO_LINE_QTY_Result.ValidationFailed Then GoTo ValidationFail
    
    ' End - Enter Line Fields
    
    
    PO_Line_ScheduleCount = purchaseOrder.PO_Lines(PO_Line).ScheduleCount
    
    For PO_Line_Schedule = 1 To PO_Line_ScheduleCount
        ' Begin - Enter Schedule Fields
        
        Dim PO_pageScheduleIndex_tmp As Integer
        PO_pageScheduleIndex_tmp = PO_pageScheduleIndex + PO_Line_Schedule - 1
        
        Debug.Print
        
        ' Due date set or PO default due date is not set
        If purchaseOrder.PO_Lines(PO_Line).Schedules(PO_Line_Schedule).ScheduleFields.DUE_DATE > 0 Or purchaseOrder.PO_Defaults.SCH_DUE_DATE = 0 Then
            PeopleSoft_Page_SetValidatedField driver:=driver, fieldElementID:=("PO_LINE_SHIP_DUE_DT$" & (PO_pageScheduleIndex + PO_Line_Schedule - 1)), _
                fieldValue:=Format(purchaseOrder.PO_Lines(PO_Line).Schedules(PO_Line_Schedule).ScheduleFields.DUE_DATE, "mm/dd/yyyy"), _
                validationResult:=purchaseOrder.PO_Lines(PO_Line).Schedules(PO_Line_Schedule).ScheduleFields.DUE_DATE_Result, _
                expectedPopupContents:="*Due Date selected is a weekend or a public holiday*"
                

            If purchaseOrder.PO_Lines(PO_Line).Schedules(PO_Line_Schedule).ScheduleFields.DUE_DATE_Result.ValidationFailed Then GoTo ValidationFail
        End If
        
    
        'Debug.Print
        PeopleSoft_Page_SetValidatedField driver, ("PO_LINE_SHIP_SHIPTO_ID$" & (PO_pageScheduleIndex + PO_Line_Schedule - 1)), _
            CStr(purchaseOrder.PO_Lines(PO_Line).Schedules(PO_Line_Schedule).ScheduleFields.SHIPTO_ID), _
            purchaseOrder.PO_Lines(PO_Line).Schedules(PO_Line_Schedule).ScheduleFields.SHIPTO_ID_Result
        
        If purchaseOrder.PO_Lines(PO_Line).Schedules(PO_Line_Schedule).ScheduleFields.SHIPTO_ID_Result.ValidationFailed Then GoTo ValidationFail
        
        
        If purchaseOrder.PO_Lines(PO_Line).Schedules(PO_Line_Schedule).ScheduleFields.QTY > 0 Then
            PeopleSoft_Page_SetValidatedField driver, ("PO_LINE_SHIP_QTY_PO$" & (PO_pageScheduleIndex + PO_Line_Schedule - 1)), _
                CStr(purchaseOrder.PO_Lines(PO_Line).Schedules(PO_Line_Schedule).ScheduleFields.QTY), _
                purchaseOrder.PO_Lines(PO_Line).Schedules(PO_Line_Schedule).ScheduleFields.QTY_Result
            
            If purchaseOrder.PO_Lines(PO_Line).Schedules(PO_Line_Schedule).ScheduleFields.QTY_Result.ValidationFailed Then
                'The vendor item price was not setup, or the corresponding UOd doesn 't meet the minimum requirements. The item standard price is used instead.
                If InStr(1, purchaseOrder.PO_Lines(PO_Line).Schedules(PO_Line_Schedule).ScheduleFields.QTY_Result.ValidationErrorText, "The item standard price is") > 0 Then
                    purchaseOrder.PO_Lines(PO_Line).Schedules(PO_Line_Schedule).ScheduleFields.QTY_Result.ValidationFailed = False
                    purchaseOrder.PO_Lines(PO_Line).Schedules(PO_Line_Schedule).ScheduleFields.QTY_Result.ValidationErrorText = ""
                End If
            End If
            
            If purchaseOrder.PO_Lines(PO_Line).Schedules(PO_Line_Schedule).ScheduleFields.QTY_Result.ValidationFailed Then GoTo ValidationFail
        End If
        
        
        ' Retrieve price Dim priceStr As String
        Dim priceStr As String, priceVal As Currency
        
        priceStr = driver.findElementById("PO_LINE_SHIP_PRICE_PO$" & (PO_pageScheduleIndex + PO_Line_Schedule - 1)).getAttribute("value")
        priceVal = CurrencyFromString(priceStr)
        
        ' Price given? Change price if PO default price is different from what is given. Otherwise, retrieve the price from page
        If purchaseOrder.PO_Lines(PO_Line).Schedules(PO_Line_Schedule).ScheduleFields.PRICE > 0 Then
            If purchaseOrder.PO_Lines(PO_Line).Schedules(PO_Line_Schedule).ScheduleFields.PRICE <> priceVal Then
            
                PeopleSoft_Page_SetValidatedField driver, _
                    ("PO_LINE_SHIP_PRICE_PO$" & (PO_pageScheduleIndex + PO_Line_Schedule - 1)), _
                    CStr(purchaseOrder.PO_Lines(PO_Line).Schedules(PO_Line_Schedule).ScheduleFields.PRICE), _
                    purchaseOrder.PO_Lines(PO_Line).Schedules(PO_Line_Schedule).ScheduleFields.PRICE_Result
                
                If purchaseOrder.PO_Lines(PO_Line).Schedules(PO_Line_Schedule).ScheduleFields.PRICE_Result.ValidationFailed Then GoTo ValidationFail
                

            End If
        Else
             purchaseOrder.PO_Lines(PO_Line).Schedules(PO_Line_Schedule).ScheduleFields.PRICE = priceVal
        End If
        ' End - Enter Schedule Fields
        
        ' Begin - Enter Distribution Fields
        
        PeopleSoft_Page_SetValidatedField driver, _
            ("BUSINESS_UNIT_PC$" & (PO_pageScheduleIndex + PO_Line_Schedule - 1)), _
            purchaseOrder.PO_Lines(PO_Line).Schedules(PO_Line_Schedule).DistributionFields.BUSINESS_UNIT_PC, _
            purchaseOrder.PO_Lines(PO_Line).Schedules(PO_Line_Schedule).DistributionFields.BUSINESS_UNIT_PC_Result
        
        If purchaseOrder.PO_Lines(PO_Line).Schedules(PO_Line_Schedule).DistributionFields.BUSINESS_UNIT_PC_Result.ValidationFailed Then GoTo ValidationFail
        
        
        PeopleSoft_Page_SetValidatedField driver, _
            ("PROJECT_ID$" & (PO_pageScheduleIndex + PO_Line_Schedule - 1)), _
            purchaseOrder.PO_Lines(PO_Line).Schedules(PO_Line_Schedule).DistributionFields.PROJECT_CODE, _
            purchaseOrder.PO_Lines(PO_Line).Schedules(PO_Line_Schedule).DistributionFields.PROJECT_CODE_Result
        
        If purchaseOrder.PO_Lines(PO_Line).Schedules(PO_Line_Schedule).DistributionFields.PROJECT_CODE_Result.ValidationFailed Then GoTo ValidationFail
        
        
        
        'PeopleSoft_Page_SetValidatedField driver, _
        '    ("ACTIVITY_ID$" & (PO_pageScheduleIndex + PO_Line_Schedule - 1)), _
        '    purchaseOrder.PO_Lines(PO_Line).Schedules(PO_Line_Schedule).DistributionFields.ACTIVITY_ID, _
        '    purchaseOrder.PO_Lines(PO_Line).Schedules(PO_Line_Schedule).DistributionFields.ACTIVITY_ID_Result
        
        'If purchaseOrder.PO_Lines(PO_Line).Schedules(PO_Line_Schedule).DistributionFields.ACTIVITY_ID_Result.ValidationFailed Then GoTo ValidationFail
        
        PeopleSoft_PurchaseOrder_SetValidatedField_ActivityID driver, _
                ("ACTIVITY_ID$" & (PO_pageScheduleIndex + PO_Line_Schedule - 1)), _
                purchaseOrder.PO_Lines(PO_Line).Schedules(PO_Line_Schedule).DistributionFields.ACTIVITY_ID, _
                purchaseOrder.PO_Lines(PO_Line).Schedules(PO_Line_Schedule).DistributionFields.ACTIVITY_ID_Result, _
                ("ACTIVITY_ID$prompt$" & (PO_pageScheduleIndex + PO_Line_Schedule - 1))
        
        If purchaseOrder.PO_Lines(PO_Line).Schedules(PO_Line_Schedule).DistributionFields.ACTIVITY_ID_Result.ValidationFailed Then GoTo ValidationFail
        
        
        If purchaseOrder.PO_Lines(PO_Line).Schedules(PO_Line_Schedule).DistributionFields.LOCATION_ID > 0 Then
            PeopleSoft_Page_SetValidatedField driver, _
                ("PO_LINE_DISTRIB_LOCATION$" & (PO_pageScheduleIndex + PO_Line_Schedule - 1)), _
                CStr(purchaseOrder.PO_Lines(PO_Line).Schedules(PO_Line_Schedule).DistributionFields.LOCATION_ID), _
                purchaseOrder.PO_Lines(PO_Line).Schedules(PO_Line_Schedule).DistributionFields.LOCATION_ID_Result
            
            If purchaseOrder.PO_Lines(PO_Line).Schedules(PO_Line_Schedule).DistributionFields.LOCATION_ID_Result.ValidationFailed Then GoTo ValidationFail
        End If
            
        ' TODO: Additional Chartfields for expenses
        '   Cost Center: DEPTID$X
        '   Product Code: PRODUCT$X
         
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


Public Function PeopleSoft_PurchaseOrder_ProcessChangeOrder(ByRef session As PeopleSoft_Session, ByRef poChangeOrder As PeopleSoft_PurchaseOrder_ChangeOrder) As Boolean
    
    
    If DEBUG_OPTIONS.CaptureExceptions Then On Error GoTo ExceptionThrown
    
    
    Dim By As New By, Assert As New Assert, Verify As New Verify
    Dim driver As New SeleniumWrapper.WebDriver
    Dim popupResult As PeopleSoft_Page_PopupCheckResult
    
    
    PeopleSoft_Login session
    
    If Not session.loggedIn Then
        poChangeOrder.GlobalError = "Logon Error: " & session.LogonError
        poChangeOrder.HasError = True
        
        PeopleSoft_PurchaseOrder_ProcessChangeOrder = False
        Exit Function
    End If

    
    Set driver = session.driver
    
    
    PeopleSoft_NavigateTo_ExistingPO session, poChangeOrder.PO_BU, poChangeOrder.PO_ID
    
    ' TODO: Check if we navigated to a PO
    If PeopleSoft_Page_ElementExists(driver, By.XPath(".//*[contains(text(),'Purchase order being processed by batch programs')]")) Then
        poChangeOrder.GlobalError = "PO currently being processed by other programs."
        poChangeOrder.HasError = True
        
        GoTo ChangeOrderFailed
    End If
    
    ' -------------------------------------------------------------------
    ' Begin - Comments Section
    ' -------------------------------------------------------------------
    If poChangeOrder.PO_HDR_FLG_SEND_TO_VENDOR <> KeepExistingValue Then
        Dim elID As String
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
    Dim PopupText As String, popUpIsExpected As Boolean
    Dim result As Boolean
    
    Dim modifyDefaults As Boolean
    
    modifyDefaults = poChangeOrder.PO_DUE_DATE > 0 Or Len(poChangeOrder.PO_PROJECT_CODE) > 0
    
    If modifyDefaults Then
        
        ' Re-use code for filling PO defaults, except only use the due date field
        Dim PO_Defaults As PeopleSoft_PurchaseOrder_PO_Defaults
        
        PO_Defaults.SCH_DUE_DATE = poChangeOrder.PO_DUE_DATE
        PO_Defaults.DIST_PROJECT_CODE = poChangeOrder.PO_PROJECT_CODE
    
        result = PeopleSoft_PurchaseOrder_PO_Defaults_Fill(driver, PO_Defaults)
        
        poChangeOrder.PO_DUE_DATE_Result = PO_Defaults.SCH_DUE_DATE_Result
        poChangeOrder.PO_PROJECT_CODE_Result = PO_Defaults.DIST_PROJECT_CODE_Result
        
        
        If result = False Then
            poChangeOrder.HasError = True
            
            If PO_Defaults.HasGlobalError Then poChangeOrder.GlobalError = PO_Defaults.GlobalError
            
            
            PeopleSoft_PurchaseOrder_ProcessChangeOrder = False
            Exit Function
        End If
        
        ' Suppress popup
        Do
            popupResult = PeopleSoft_Page_CheckForPopup(driver, acknowledgePopup:=True, raiseErrorIfUnexpected:=False, _
                            expectedContent:=Array( _
                                "*Default values will be applied only to PO lines that are not received or invoiced*", _
                                "*This action will create a change order*", _
                                "*This PO has been dispatched, add/delete/change a line or schedule will create a change order*" _
                                ))
            
            If popupResult.IsExpected = False Then
                poChangeOrder.HasError = True
                poChangeOrder.GlobalError = "Unexpected Popup: " & popupResult.PopupText
        
                GoTo ChangeOrderFailed
            End If
        Loop While popupResult.HasPopup
        
        PopupText = PeopleSoft_Page_SuppressPopup(driver, vbOK)
        
        'If Len(popUpText) > 0 Then
        '    popUpIsExpected = InStr(1, popUpText, "Default values will be applied only to PO lines that are not received or invoiced") > 0
        '    PeopleSoft_Page_WaitForProcessing driver
        '
        '    If popUpIsExpected = False Then
        '        poChangeOrder.HasError = True
        '        poChangeOrder.GlobalError = "Unexpected Popup: " & popUpText
        '
        '        GoTo ChangeOrderFailed
        '    End If
        'End If
        
        'driver.Wait 500 ' wait 0.5s
        
        PopupText = PeopleSoft_Page_SuppressPopup(driver, vbOK)
        
        'If Len(popUpText) > 0 Then
        '    popUpIsExpected = InStr(1, popUpText, "This action will create a change order") > 0
        '
        '    If popUpIsExpected = False Then
        '        poChangeOrder.HasError = True
        '        poChangeOrder.GlobalError = "Unexpected Popup: " & popUpText
        '
        '        GoTo ChangeOrderFailed
        '    End If
        'End If
        PopupText = PeopleSoft_Page_SuppressPopup(driver, vbOK)
        
        'If Len(popUpText) > 0 Then
        '    popUpIsExpected = InStr(1, popUpText, "This PO has been dispatched, add/delete/change a line or schedule will create a change order.") > 0
        '
        '    If popUpIsExpected = False Then
        '        poChangeOrder.HasError = True
        '        poChangeOrder.GlobalError = "Unexpected Popup: " & popUpText
        '
        '        GoTo ChangeOrderFailed
        '    End If
        'End If
        
        ' If we modified the defaults, then collapse all - required for any individual change o        driver.runScript "javascript:submitAction_win0(document.win0,'PO_PNLS_PB_COLLAPSE_ALL_PB', 0, 0, 'Collapse All', false, true);" ' Fix for 2.9.1.1  due to PS upgrade
        'driver.runScript "javascript:hAction_win0(document.win0,'PO_PNLS_PB_COLLAPSE_ALL_PB', 0, 0, 'Collapse All', false, true);"
        PeopleSoft_Page_WaitForProcessing driver
        
        ' driver.runScript "javascript:hAction_win0(document.win0,'PO_PNLS_PB_EXPAND_ALL_PB', 0, 0, 'Expand All', false, true);"
    End If
    ' -------------------------------------------------------------------
    ' End - PO Defaults Section
    ' -------------------------------------------------------------------
    


    
    ' ---------------------------------------------------------------
    ' Begin - Process Change Orders for individual items/schedules
    ' ---------------------------------------------------------------
    Dim paginationText As String, posTo As Integer, posOf As Integer
    Dim pageLineFrom As Integer, pageLineTo As Integer, pageLineTotal As Integer
    Dim anyLineEditsOnPage As Boolean
    
    Dim isSinglePagePO As Boolean
    
    Dim numProcessed As Integer
    
    Dim i As Integer
    
   
    If poChangeOrder.PO_ChangeOrder_ItemCount > 0 Then
        
        ' The below
        
        isSinglePagePO = True
        numProcessed = 0
        
        Do
            anyLineEditsOnPage = False
            
            
            If PeopleSoft_Page_ElementExists(driver, By.id("PO_SCR_NAV_WRK_SRCH_RSLT_MSG")) Then
                isSinglePagePO = False
            
                paginationText = driver.findElementById("PO_SCR_NAV_WRK_SRCH_RSLT_MSG").Text  ' example: 1 to 75 of 77
            
                posTo = InStr(1, paginationText, " to ")
                posOf = InStr(1, paginationText, " of ")
                
                Debug.Assert posTo > 0
                Debug.Assert posOf > 0
                Debug.Assert posOf > posTo
                
                pageLineFrom = Mid(paginationText, 1, posTo - 1)
                pageLineTo = Mid(paginationText, posTo + Len(" to "), posOf - posTo - Len(" to "))
                pageLineTotal = Mid(paginationText, posOf + Len(" of "))
                
                anyLineEditsOnPage = False
                
                For i = 1 To poChangeOrder.PO_ChangeOrder_ItemCount
                    If pageLineFrom <= poChangeOrder.PO_ChangeOrder_Items(i).PO_Line And poChangeOrder.PO_ChangeOrder_Items(i).PO_Line <= pageLineTo Then
                        anyLineEditsOnPage = True
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
            ' This entire section can be removed after the issue is fixed (not likely)
            If True Then
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
            
            If anyLineEditsOnPage Or isSinglePagePO Then
                Dim pageLineIndex As Integer, pageSchIndex As Integer
                Dim lineIndex As Integer
                
                pageLineIndex = 0
                pageSchIndex = 0
                
                
                
                ' Expand All
                driver.runScript "javascript:submitAction_win0(document.win0,'PO_PNLS_PB_EXPAND_ALL_PB', 0, 0, 'Expand All', false, true);" ' Fix for 2.9.1.1  due to PS upgrade
                'driver.runScript "javascript:hAction_win0(document.win0,'PO_PNLS_PB_EXPAND_ALL_PB', 0, 0, 'Expand All', false, true);"
                PeopleSoft_Page_WaitForProcessing driver
                
                For i = 1 To poChangeOrder.PO_ChangeOrder_ItemCount
                    If pageLineFrom <= poChangeOrder.PO_ChangeOrder_Items(i).PO_Line And poChangeOrder.PO_ChangeOrder_Items(i).PO_Line <= pageLineTo Then
                        lineIndex = poChangeOrder.PO_ChangeOrder_Items(i).PO_Line - pageLineFrom
                        
                        
                        ' TODO: Check if PO_LINE_CANCEL_STATUS$1 == Active
                        
                        ' Determine the schedule index in the page by looking at the index for the schedule captions
                        Dim webElemScheduleCaptions As SeleniumWrapper.WebElementCollection, webElem As SeleniumWrapper.WebElement
                        Dim webElemScheduleCaptionId As String
                        Set webElemScheduleCaptions = driver.findElementsByXPath(".//*[@id='ACE_PO_LINE_SHIP_SCROL$" & lineIndex & "']/descendant::*[contains(@id,'win0divPO_LINE_SHIP_SCHED_NBR')]/span")
                        
                        pageSchIndex = -1
                        
                        For Each webElem In webElemScheduleCaptions
                            If CInt(webElem.Text) = poChangeOrder.PO_ChangeOrder_Items(i).PO_Schedule Then
                                ' Extract schedule index from span ID (Example: PO_LINE_SHIP_SCHED_NBR$10)
                                webElemScheduleCaptionId = webElem.getAttribute("id")
                                pageSchIndex = CInt(Mid$(webElemScheduleCaptionId, InStr(1, webElemScheduleCaptionId, "$") + 1))
                                Exit For
                            End If
                        Next webElem
                        
                        ' TODO: Check if pageSchIndex >= 0
                        
                        
                        ' Expand Schedule
                        'driver.runScript "javascript:hAction_win0(document.win0,'PO_PNLS_PB_EXPAND_PB$" & lineIndex & "', 0, 0, 'Expand Schedule Section', false, true);"
                        'PeopleSoft_Page_WaitForProcessing driver
                        
                        ' Expand Distribution
                        ' Click PO_PNLS_PB_EXPAND_PB$232$$0
                         'driver.runScript "javascript:hAction_win0(document.win0,'PO_PNLS_PB_EXPAND_PB$232$$0', 0, 0, 'Expand Distribution Section', false, true);"
                        'PeopleSoft_Page_WaitForProcessing driver
                        
                        
                        ' Click PO_PNLS_WRK_CHANGE_SHIP$0 - TODO: Check if it exists
                        driver.runScript "javascript:submitAction_win0(document.win0,'PO_PNLS_WRK_CHANGE_SHIP$" & pageSchIndex & "');" ' Fix for 2.9.1.1  due to PS upgrade
                        'driver.runScript "javascript:hAction_win0(document.win0,'PO_PNLS_WRK_CHANGE_SHIP$" & pageSchIndex & "');"
                        PeopleSoft_Page_WaitForProcessing driver
                        
                        
                        driver.runScript "javascript:submitAction_win0(document.win0,'PO_PNLS_PB_EXPAND_ALL_PB', 0, 0, 'Expand All', false, true);" ' Fix for 2.9.1.1  due to PS upgrade
        
                        
                        ' Note since 2.9.1.1,
                        '<a id="PO_PNLS_WRK_Z_CHANGE_DIST$0" href="javascript:submitAction_win0(document.win0,'PO_PNLS_WRK_Z_CHANGE_DIST$0', false, true);" tabindex="893" name="PO_PNLS_WRK_Z_CHANGE_DIST$0">
                        'a id="PO_PNLS_WRK_GOTO_SCHED_DTLS$0" href="javascript:submitAction_win0(document.win0,'PO_PNLS_WRK_GOTO_SCHED_DTLS$0');" tabindex="584" name="PO_PNLS_WRK_GOTO_SCHED_DTLS$0">
                        '<a id="PO_PNLS_WRK_GOTO_LINE_DTLS$2" href="javascript:submitAction_win0(document.win0,'PO_PNLS_WRK_GOTO_LINE_DTLS$2');" tabindex="557" name="PO_PNLS_WRK_GOTO_LINE_DTLS$2">
                        
                       ' <a href="javascript:hAction_win0(document.win0,'PO_PNLS_WRK_CHANGE_LINE', 0, 0, 'Create Line Change', false, true);" tabindex="16" id="PO_PNLS_WRK_CHANGE_LINE" name="PO_PNLS_WRK_CHANGE_LINE"><img border="0" title="Create Line Change" alt="Create Line Change" name="PO_PNLS_WRK_CHANGE_LINE$IMG" src="/cs/ps/cache/PS_DELTA_ICN_1.gif"></a>
                        Dim tmp As String
                        tmp = driver.findElementById("PO_LINE_DUE_DATE$" & (pageSchIndex)).getAttribute("disabled")
                        
                        If poChangeOrder.PO_ChangeOrder_Items(i).SCH_DUE_DATE > 0 Then
                            PeopleSoft_Page_SetValidatedField driver, _
                                ("PO_LINE_DUE_DATE$" & (pageSchIndex)), _
                                Format(poChangeOrder.PO_ChangeOrder_Items(i).SCH_DUE_DATE, "mm/dd/yyyy"), _
                                poChangeOrder.PO_ChangeOrder_Items(i).SCH_DUE_DATE_Result
                                
                            If poChangeOrder.PO_ChangeOrder_Items(i).SCH_DUE_DATE_Result.ValidationFailed Then GoTo ValidationFail
                        End If
                        
                        numProcessed = numProcessed + 1
                    End If
                Next i
                
                
            End If
            
            Debug.Print
            
            If pageLineTo < pageLineTotal And numProcessed < poChangeOrder.PO_ChangeOrder_ItemCount And Not isSinglePagePO Then
                ' Next page
                driver.findElementById("PO_SCR_NAV_WRK_NEXT_ITEM_BUTTON").Click
                PeopleSoft_Page_WaitForProcessing driver
                
                'driver.runScript "javascript:hAction_win0(document.win0,'PO_SCR_NAV_WRK_NEXT_ITEM_BUTTON', 0, 0, '', false, true);"
                
                
                PopupText = PeopleSoft_Page_SuppressPopup(driver, vbOK)
                
                PeopleSoft_Page_WaitForProcessing driver, TIMEOUT_LONG
            End If
        Loop Until pageLineTo = pageLineTotal Or isSinglePagePO
        
    
    End If
    ' ---------------------------------------------------------------
    ' End - Process Change Orders for individual items/schedules
    ' ---------------------------------------------------------------
    
    
    
    
    driver.findElementById("PO_KK_WRK_PB_BUDGET_CHECK").Click
    
    PeopleSoft_Page_WaitForProcessing driver, TIMEOUT_LONG
    
    ' TODO: Check if change made (e.g., due date was actually changed)
    
    ' Change to: <span class="PATRANSACTIONTITLE">Change Reason</span>
    If PeopleSoft_Page_ElementExists(driver, By.id("PO_CHNG_REASON_COMMENTS$0")) Then
        
        driver.findElementById("PO_CHNG_REASON_COMMENTS$0").Clear
        driver.findElementById("PO_CHNG_REASON_COMMENTS$0").SendKeys poChangeOrder.ChangeReason
        
        
        
        driver.findElementById("#ICSave").Click
        PeopleSoft_Page_WaitForProcessing driver, TIMEOUT_LONG
    End If
    
    PeopleSoft_PurchaseOrder_ProcessChangeOrder = True
    Exit Function
    
ValidationFail:
    poChangeOrder.HasError = True
    
ChangeOrderFailed:
    PeopleSoft_PurchaseOrder_ProcessChangeOrder = False
    Exit Function
    
ExceptionThrown:
    
    poChangeOrder.HasError = True
    poChangeOrder.GlobalError = "Exception: " & Err.Description
    
    PeopleSoft_PurchaseOrder_ProcessChangeOrder = False

End Function


Public Function PeopleSoft_Receipt_CreateReceipt(ByRef session As PeopleSoft_Session, ByRef rcpt As PeopleSoft_Receipt) As Boolean


    If DEBUG_OPTIONS.CaptureExceptions Then On Error GoTo ExceptionThrown

    'Dim session As PeopleSoft_Session
    Dim driver As New SeleniumWrapper.WebDriver
    Dim elem As WebElement
    
    
    Set driver = session.driver
    
    
    Dim By As New By, Assert As New Assert, Verify As New Verify
    Dim i As Integer, j As Integer
    
    ' Pre-Check ensure there are no duplicate PO lines/schedules in ReceiptItems
    For i = 1 To rcpt.ReceiptItemCount
        For j = i + 1 To rcpt.ReceiptItemCount
            If rcpt.ReceiptItems(i).PO_Line = rcpt.ReceiptItems(j).PO_Line And rcpt.ReceiptItems(i).PO_Schedule = rcpt.ReceiptItems(j).PO_Schedule Then
                rcpt.ReceiptItems(i).HasError = True
                rcpt.ReceiptItems(j).HasError = True
                rcpt.ReceiptItems(i).ItemError = "Duplicate line/schedule"
                rcpt.ReceiptItems(j).ItemError = "Duplicate line/schedule"
                
                rcpt.HasGlobalError = True
                rcpt.GlobalError = "Duplicate line/schedule in one more receipt lines"
            End If
        Next j
    Next i
    
    If rcpt.HasGlobalError = True Then GoTo ReceiptFailed
    
    
    
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
        anyItemHasErrors = False
        
        For i = 1 To rcpt.ReceiptItemCount
            pageRcptLineIdx = mapReceiptItemsToPageReceiptLines(i)
            
            If rcpt.ReceiptItems(i).HasError = False And rcpt.ReceiptItems(i).IsNotReceivable = False Then
                rcpt.ReceiptItems(i).Accept_Qty = pageReceiptLines(pageRcptLineIdx).Accept_Qty
                        
                ' Check: receive quantity is less than accept qty
                If rcpt.ReceiptItems(i).Receive_Qty > 0 Then ' Receipt qty specified
                    If rcpt.ReceiptItems(i).Receive_Qty > rcpt.ReceiptItems(i).Accept_Qty Then
                        rcpt.ReceiptItems(i).HasError = True
                        rcpt.ReceiptItems(i).ItemError = "Receive qty is greater than accepted qty (Accept Qty: " & rcpt.ReceiptItems(i).Accept_Qty & ")"
                        anyItemHasErrors = True
                    End If
                End If
                        
                ' Second check passed
                If rcpt.ReceiptItems(i).HasError = False Then
                    If rcpt.ReceiptItems(i).Receive_Qty > 0 Then ' Receipt qty specified - otherwise receive all
                        Dim tmpValidationResult As PeopleSoft_Field_ValidationResult
                        
                        PeopleSoft_Page_SetValidatedField driver, ("RECV_LN_SHIP_QTY_SH_RECVD$" & pageReceiptLines(pageRcptLineIdx).PageTableRowIndex), _
                            CStr(rcpt.ReceiptItems(i).Receive_Qty), tmpValidationResult
                        
                        If tmpValidationResult.ValidationFailed Then
                            rcpt.ReceiptItems(i).HasError = True
                            rcpt.ReceiptItems(i).ItemError = "RECEIVE QTY ERROR: " & tmpValidationResult.ValidationErrorText
                            anyItemHasErrors = True
                        End If
                    Else
                        ' No receipt qty give. Receive on all and return the qty.
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
        
            
    ' Save Receipt
    'driver.findElementById("#ICSave").Click
    driver.runScript "javascript:setSaveText_win0('Saving...');submitAction_win0(document.win0, '#ICSave');"
    
    
    ' Wait for "Saving..." to stop.
    driver.waitForElementPresent "css=#SAVED_win0"
    'driver.findElementById("processing").waitForCssValue "visibility", "visible"
    driver.findElementById("SAVED_win0").waitForCssValue "visibility", "hidden"
    
    
     
    Dim popupCheckResult As PeopleSoft_Page_PopupCheckResult
    
    Debug.Print "Expecting Popup: Have these receipt quantities been checked for accuracy"
    
    popupCheckResult = PeopleSoft_Page_CheckForPopup(driver:=driver, acknowledgePopup:=False, _
        raiseErrorIfUnexpected:=False, expectedContent:="*Have these receipt quantities been checked for accuracy*")
    
    
    If popupCheckResult.HasPopup = False Or (popupCheckResult.HasPopup And popupCheckResult.IsExpected = False) Then
        rcpt.HasGlobalError = True
        rcpt.GlobalError = "Did not receive expected popup: Have these receipt quantities been checked for accuracy?" _
                            & IIf(popupCheckResult.HasPopup, vbCrLf & "Popup received: " & popupCheckResult.PopupText, "")
        
        GoTo ReceiptFailed
    End If
    
    ' We received correct popup -> acknowledge
    PeopleSoft_Page_AcknowledgePopup driver, popupCheckResult, vbYes
    PeopleSoft_Page_WaitForProcessing driver
    
            
            
    


    
    
    ' Check for receipt ID.
    rcpt.RECEIPT_ID = driver.findElementById("RECV_HDR_RECEIVER_ID").Text
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
            Debug.Print "Popup received after Receipt " & popupCountCheck & ": " & popupCheckResult.PopupText
            'rcpt.GlobalError = rcpt.GlobalError & "Popup Received after Receipt " & popupCountCheck & ": " & popupCheckResult.PopupText & vbCrLf
        
            PeopleSoft_Page_WaitForProcessing driver
        End If
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
    Exit Function
       
ExceptionThrown:
    PeopleSoft_Receipt_CreateReceipt = False
    
    rcpt.HasGlobalError = True
    rcpt.GlobalError = "Exception: " & Err.Description
    



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
    
    For rowIndex = 0 To numReturnedRows - 1
        Dim Row_CheckDisabled As String
      
        unreceivedItems(rowIndex + 1).PageTableRowIndex = rowIndex
      
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
    
    

End Function
' PeopleSoft_Receipt_ExtractReceiptItemsFromPage: Extracts all receipt items from PS Receipt page. Assumes we already navigated to page. Returns count but populated the paremeter pageReceiptItems
Private Function PeopleSoft_Receipt_ExtractReceiptLinesFromPage(driver As SeleniumWrapper.WebDriver, ByRef pageReceiptLines() As PeopleSoft_ReceiptPage_ReceiptLine) As Long


    Dim pageReceiptLineCount As Long
    
    
    pageReceiptLineCount = 0
    PeopleSoft_Receipt_ExtractReceiptLinesFromPage = 0
    
    
    ' Simulate "View All"
    driver.runScript "javascript:submitAction_win0(document.win0,'RECV_LN_SHIP$hviewall$0');"
    PeopleSoft_Page_WaitForProcessing driver
    
    
    pageReceiptLineCount = driver.getXpathCount(".//*[contains(@id,'ftrRECV_LN_SHIP$0_row')]")
    If pageReceiptLineCount = 0 Then Exit Function
    
    
    ' Simulate "Show All Columns"
    driver.runScript "javascript:submitAction_win0(document.win0,'RECV_LN_SHIP$htab$0');"
    PeopleSoft_Page_WaitForProcessing driver
    

    ReDim pageReceiptLines(1 To pageReceiptLineCount) As PeopleSoft_ReceiptPage_ReceiptLine
    
    
    Dim rowIndex As Integer
    
    
    ' Note: We need to use javascript to get the element content/values since nothing is returned when items are not visible on the page
    For rowIndex = 0 To pageReceiptLineCount - 1
        pageReceiptLines(rowIndex + 1).PageTableRowIndex = rowIndex
        
        
        pageReceiptLines(rowIndex + 1).Receipt_Line = CInt(driver.executeScript("javascript: return document.getElementById('RECV_LN_NBR$" & rowIndex & "').textContent;"))
        
        pageReceiptLines(rowIndex + 1).Item_ID = driver.executeScript("javascript: return document.getElementById('INV_ITEM_ID$" & rowIndex & "').textContent;")
        pageReceiptLines(rowIndex + 1).Description = driver.executeScript("javascript: return document.getElementById('DESCR$" & rowIndex & "').textContent;")
        pageReceiptLines(rowIndex + 1).Status = driver.executeScript("javascript: return document.getElementById('RECV_LN_SHIP_RECV_SHIP_STATUS$" & rowIndex & "').textContent;")
        
        pageReceiptLines(rowIndex + 1).Receipt_Price = CDec(driver.executeScript("javascript: return document.getElementById('PRICE_RECV$" & rowIndex & "').textContent;"))
        pageReceiptLines(rowIndex + 1).Accept_Qty = CDec(driver.executeScript("javascript: return document.getElementById('RECV_LN_SHIP_QTY_SH_ACCPT$" & rowIndex & "').textContent;"))
        
        pageReceiptLines(rowIndex + 1).PO_Line = CLng(driver.executeScript("javascript: return document.getElementById('RECV_LN_SHIP_LINE_NBR$" & rowIndex & "').textContent;"))
        pageReceiptLines(rowIndex + 1).PO_Schedule = CLng(driver.executeScript("javascript: return document.getElementById('RECV_LN_SHIP_SCHED_NBR$" & rowIndex & "').textContent;"))
        
        pageReceiptLines(rowIndex + 1).Receipt_Qty = CDec(driver.executeScript("javascript: return document.getElementById('RECV_LN_SHIP_QTY_SH_RECVD$" & rowIndex & "').value;"))
    Next rowIndex
    
    
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
    
    
    On Error GoTo ExceptionThrown
    
    
    Dim By As New By, Assert As New Assert, Verify As New Verify
    Dim driver As SeleniumWrapper.WebDriver
    
    
    PeopleSoft_Login session
    
    If Not session.loggedIn Then
        poRetryBC.GlobalError = "Logon Error: " & session.LogonError
        poRetryBC.HasGlobalError = True
        
        GoTo RetryBCFailed
    End If

    
    Set driver = session.driver
    
    
    PeopleSoft_NavigateTo_ExistingPO session, poRetryBC.PO_BU, poRetryBC.PO_ID
    
    ' TODO: Check if we navigated to a PO
    
    If PeopleSoft_Page_ElementExists(driver, By.XPath(".//*[text()='PO Budget Check Errors']")) Then
        driver.findElementById("#ICCancel").Click
        'driver.runScript "javascript:submitAction_win0(document.win0, '#ICCancel');"
        
        PeopleSoft_Page_WaitForProcessing driver
    End If
    
    
    ' Skip if PO is Dispatched or Approved.
    Dim poStatusText As String
    poStatusText = driver.findElementById("PSXLATITEM_XLATSHORTNAME").Text
    
    If poStatusText = "Approved" Or poStatusText = "Dispatched" Then
        PeopleSoft_PurchaseOrder_RetrySaveWithBudgetCheck = True
        Exit Function
    End If
    
    
    Dim result As Boolean
    
    result = PeopleSoft_PurchaseOrder_SaveWithBudgetCheck(driver, poRetryBC.BudgetCheck_Result)
    
    If result = False Then
        poRetryBC.GlobalError = poRetryBC.BudgetCheck_Result.GlobalError
        poRetryBC.HasGlobalError = poRetryBC.BudgetCheck_Result.HasGlobalError
        
        GoTo RetryBCFailed
    End If
    
    
    PeopleSoft_PurchaseOrder_RetrySaveWithBudgetCheck = True
    Exit Function
    
    
ValidationFailed:
RetryBCFailed:
    PeopleSoft_PurchaseOrder_RetrySaveWithBudgetCheck = False
    Exit Function
       
ExceptionThrown:
    PeopleSoft_PurchaseOrder_RetrySaveWithBudgetCheck = False
    
    poRetryBC.HasGlobalError = True
    poRetryBC.GlobalError = "Exception: " & Err.Description
    


End Function



Public Function PeopleSoft_Page_SetValidatedField(ByRef driver As SeleniumWrapper.WebDriver, ByVal fieldElementID As String, ByVal fieldValue As String, ByRef validationResult As PeopleSoft_Field_ValidationResult, Optional ignoreEmptyValues As Boolean = True, Optional expectedPopupContents As Variant) As Boolean


     
    
    validationResult.ValidationFailed = False
    validationResult.ValidationErrorText = ""

    
        
    
    ' Dont bother if value is empty string or option to ignore empty values is false
    If Len(fieldValue) > 0 Or ignoreEmptyValues = False Then
        Dim elID As String, elVal As String
        
        'elID = fieldElement.getAttribute("id")
        elID = Replace(fieldElementID, "'", "\'")
        
        elVal = driver.executeScript("return document.getElementById('" & elID & "').value;")
        
        'Dim tryNo As Integer
    
        'tryNo = 1
        
        
        'Do

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
            'driver.runScript "javascript:oChange_win0=document.getElementById('" & elID & "');addchg_win0(oChange_win0);doFocus_win0(addchg_win0, true, true);"
            'driver.runScript "javascript:addchg_win0(document.getElementById('" & elID & "'));oChange_win0=document.getElementById('" & elID & "');submitAction_win0(oChange_win0.form,oChange_win0.name);"
    
            PeopleSoft_Page_WaitForProcessing driver
            
            
            Dim popupResult As PeopleSoft_Page_PopupCheckResult
            
            driver.setImplicitWait 100 ' new in 2.11: override implicit wait (speeds up field entering)
            
            popupResult = PeopleSoft_Page_CheckForPopup(driver:=driver, acknowledgePopup:=True, raiseErrorIfUnexpected:=False, expectedContent:=expectedPopupContents)
            
            driver.setImplicitWait TIMEOUT_IMPLICIT_WAIT ' new in 2.11: restore implicit wait
            
            If popupResult.HasPopup And popupResult.IsExpected = False Then
                validationResult.ValidationErrorText = popupResult.PopupText
                validationResult.ValidationFailed = True
                
                PeopleSoft_Page_SetValidatedField = False
                Exit Function
            End If
            
            
            'fieldValResult.ValidationErrorText = PeopleSoft_Page_SuppressPopup(driver, vbOK)
            'fieldValResult.ValidationFailed = fieldValResult.ValidationErrorText <> ""
            
            
            
            'driver.Wait 500
    
            'elVal = driver.executeScript("return document.getElementById('" & elID & "').value;")
            
            
            'If tryNo >= 3 Then
            '    validationResult.ValidationFailed = True
            '    validationResult.ValidationErrorText = "Could not set value"
            'End If
            
            
            
        
        End If
    
            'tryNo = tryNo + 1
        'Loop Until elVal <> ""

    End If
   
   
    
    PeopleSoft_Page_SetValidatedField = True
    Exit Function
    
    'PeopleSoft_Page_SetValidatedField = Not fieldValResult.ValidationFailed

End Function
' Utility function to create the PO data structure in one line. Must use PeopleSoft_PurchaseOrder_AddLineSimple() to add lines
Public Function PeopleSoft_PurchaseOrder_NewPO(poBU As String, buyerID As Long, vendor As String, poReference As String) As PeopleSoft_PurchaseOrder

    Dim po As PeopleSoft_PurchaseOrder
    
    po.PO_Fields.PO_BUSINESS_UNIT = poBU
    po.PO_Fields.PO_HDR_BUYER_ID = buyerID
    po.PO_Fields.PO_HDR_PO_REF = poReference
    
    If IsNumeric(vendor) Then
        po.PO_Fields.PO_HDR_VENDOR_ID = CLng(vendor)
    Else
        po.PO_Fields.VENDOR_NAME_SHORT = vendor
    End If

End Function
' Utility function to add a line to a PO structure with a single schedule.
Public Sub PeopleSoft_PurchaseOrder_AddLineSimple(ByRef purchaseOrder As PeopleSoft_PurchaseOrder, lineItemID As String, lineItemDesc As String, schQty As Variant, shipDueDate As Date, shipToId As Long, distBusinessUnit As String, distProjectCode As String, distActivityID As String, Optional locationID As Long = 0, Optional schPrice As Currency = 0)

    
    Dim PO_LineCount As Integer
    
    
    PO_LineCount = purchaseOrder.PO_LineCount + 1

    ReDim Preserve purchaseOrder.PO_Lines(1 To PO_LineCount) As PeopleSoft_PurchaseOrder_Line
    
    ReDim purchaseOrder.PO_Lines(PO_LineCount).Schedules(1 To 1) As PeopleSoft_PurchaseOrder_Schedule
    
    purchaseOrder.PO_Lines(PO_LineCount).ScheduleCount = 1
    
    purchaseOrder.PO_Lines(PO_LineCount).LineFields.PO_LINE_ITEM_ID = lineItemID
    purchaseOrder.PO_Lines(PO_LineCount).LineFields.PO_LINE_DESC = lineItemDesc
    'purchaseOrder.PO_Lines(PO_LineCount).LineFields.PO_LINE_QTY = lineQty
    
    purchaseOrder.PO_Lines(PO_LineCount).Schedules(1).ScheduleFields.DUE_DATE = shipDueDate
    purchaseOrder.PO_Lines(PO_LineCount).Schedules(1).ScheduleFields.SHIPTO_ID = shipToId
    purchaseOrder.PO_Lines(PO_LineCount).Schedules(1).ScheduleFields.QTY = CDec(schQty)
    purchaseOrder.PO_Lines(PO_LineCount).Schedules(1).ScheduleFields.PRICE = schPrice
    purchaseOrder.PO_Lines(PO_LineCount).Schedules(1).DistributionFields.BUSINESS_UNIT_PC = distBusinessUnit
    purchaseOrder.PO_Lines(PO_LineCount).Schedules(1).DistributionFields.PROJECT_CODE = distProjectCode
    purchaseOrder.PO_Lines(PO_LineCount).Schedules(1).DistributionFields.ACTIVITY_ID = distActivityID
    purchaseOrder.PO_Lines(PO_LineCount).Schedules(1).DistributionFields.LOCATION_ID = locationID
    
    purchaseOrder.PO_LineCount = PO_LineCount
    
End Sub
Public Function PeopleSoft_PurchaseOrder_SaveWithBudgetCheck(driver As SeleniumWrapper.WebDriver, ByRef budgetCheckResult As PeopleSoft_PurchaseOrder_BudgetCheckResult) As Boolean

    If DEBUG_OPTIONS.AddFunctionNameToExceptions Then On Error GoTo ExceptionThrown


    ' ---------------------------------------------------------------------
    ' Begin - Save w/ Budget Check
    ' ---------------------------------------------------------------------
    
    Dim By As New SeleniumWrapper.By
    Dim popupResult As PeopleSoft_Page_PopupCheckResult
    
    
    
    Dim swByPOId As SeleniumWrapper.By
    Dim wePOId As SeleniumWrapper.WebElement
    
    
    driver.findElementById("PO_KK_WRK_PB_BUDGET_CHECK").Click
    PeopleSoft_Page_WaitForProcessing driver, TIMEOUT_LONG * 2 ' increase long timeout by a factor of 2 due to PS responsiveness
    
    ' Acknowledge Popup with message: The below PO line schedules exist with $0.00 or blank pricing.
    'if  below PO line schedules exist with $0.00 or blank pricing
    popupResult = PeopleSoft_Page_CheckForPopup(driver:=driver, acknowledgePopup:=True, expectedContent:="*below PO line schedules exist with $0.00 or blank pricing*")
    
    If popupResult.HasPopup And popupResult.IsExpected = False Then
        budgetCheckResult.GlobalError = "Unexpected Popup: PopupText"
        budgetCheckResult.HasGlobalError = True
        
        PeopleSoft_PurchaseOrder_SaveWithBudgetCheck = False
        Exit Function
    End If
    
    ' Begin - Deal with the new screen which asks about quantities in available excess
    If PeopleSoft_Page_ElementExists(driver, By.XPath(".//title[contains(text(),'Excess Available')]")) Then
        'elemExists = PeopleSoft_Page_ElementExists(driver, By.XPath(".//*[contains(text(),'Excess Available')]"))
        'elemExists = PeopleSoft_Page_ElementExists(driver, By.XPath(".//*[contains(text(),'The following equipment currently exists in Available Excess Inventory')]"))
        'elemExists = PeopleSoft_Page_ElementExists(driver, By.id("Z_CAT_AVL_WRK_IGNORE_PB"))
    
        driver.findElementById("Z_CAT_AVL_WRK_IGNORE_PB").Click
        driver.runScript "javascript: submitAction_win0(document.win0,'Z_CAT_AVL_WRK_IGNORE_PB');"
        PeopleSoft_Page_WaitForProcessing driver, TIMEOUT_LONG ' increase long timeout by a factor of 2 due to PS responsiveness
    End If
    ' End - Deal with the new screen which asks about quantities in available excess
    
    
    
    ' Adaptation from Oscar Gonzale code
    Dim field_result As PeopleSoft_Field_ValidationResult
    
    'driver.runScript "oParentWin.submitAction_win0(oParentWin.document.win0, '#ICOK');closeMsg(null,modId);" ' i think this clicks a popup???
    'PeopleSoft_Page_WaitForProcessing driver
    'PeopleSoft_Page_SetValidatedField driver, "Z_PO_RSN_CD_REASON_CD$0", "OVERAGE-3", field_result
    'PeopleSoft_Page_WaitForProcessing driver
    'PeopleSoft_Page_SetValidatedField driver, "Z_PO_RSN_CD_COMMENTS_254$0", "Price Change After Quote Issued", field_result
    'PeopleSoft_Page_WaitForProcessing driver
    'driver.findElementById("PO_PNLS_WRK_PB_OK$0").Click
    'PeopleSoft_Page_WaitForProcessing driver
    'driver.findElementById("Z_CAT_AVL_WRK_IGNORE_PB").Click
    'PeopleSoft_Page_WaitForProcessing driver
    'driver.findElementById("#ICCancel").Click
    'PeopleSoft_Page_WaitForProcessing driver
    'driver.findElementById("Z_E_QT_WRK_APPLY").Click
    'PeopleSoft_Page_WaitForProcessing driver
    
    
    popupResult = PeopleSoft_Page_CheckForPopup(driver:=driver, acknowledgePopup:=True)
    
    If popupResult.HasPopup Then ' Error while saving
        budgetCheckResult.GlobalError = popupResult.PopupText
        budgetCheckResult.HasGlobalError = True
        
        PeopleSoft_PurchaseOrder_SaveWithBudgetCheck = False
        Exit Function
    End If
    
    
    
    
    ' The PO ID will show up in one of two elements:
    '     Budget Check Failed -> Z_KK_ERR_WRK_PO_ID
    '     Budget Check Pass -> PO_HDR_PO_ID$14$
    '
    ' We will check for both. In some cases, neither is available immediately
    ' so we need try a few times or error out.
    Set swByPOId = By.id("Z_KK_ERR_WRK_PO_ID")
    
    Dim elementExists_PO_ID_budgetCheckFailed As Boolean  ' Case: budget check failed
    Dim elementExists_PO_ID_budgetCheckPassed As Boolean  ' Case: budget check passed
    Dim tryNo As Integer
    
    Do
        tryNo = 0
        
        elementExists_PO_ID_budgetCheckFailed = False
        elementExists_PO_ID_budgetCheckPassed = False
    
        elementExists_PO_ID_budgetCheckFailed = PeopleSoft_Page_ElementExists(driver, By.id("Z_KK_ERR_WRK_PO_ID"))
        If elementExists_PO_ID_budgetCheckFailed = False Then elementExists_PO_ID_budgetCheckPassed = PeopleSoft_Page_ElementExists(driver, By.id("PO_HDR_PO_ID$14$"))
        
        tryNo = tryNo + 1
    Loop Until elementExists_PO_ID_budgetCheckFailed Or elementExists_PO_ID_budgetCheckPassed Or tryNo = 3
    
    If elementExists_PO_ID_budgetCheckFailed = False And elementExists_PO_ID_budgetCheckPassed = False Then
        budgetCheckResult.GlobalError = "Could not find PO ID on page: manual check required"
        budgetCheckResult.HasGlobalError = True
        
        PeopleSoft_PurchaseOrder_SaveWithBudgetCheck = False
    End If
    
    If elementExists_PO_ID_budgetCheckPassed Then
        ' Budget check passed
        Set wePOId = driver.findElementById("PO_HDR_PO_ID$14$") ' Fix for 2.9.1.1
        
        If wePOId.Text = "NEXT" Then ' Error while saving
            budgetCheckResult.GlobalError = "Unknown error - Invalid PO ID Generated: " & wePOId.Text
            budgetCheckResult.HasGlobalError = True
            
            PeopleSoft_PurchaseOrder_SaveWithBudgetCheck = False
            Exit Function
        Else
            budgetCheckResult.PO_ID = wePOId.Text
        End If
    Else
        ' Budget check failed
        Set wePOId = driver.findElementById("Z_KK_ERR_WRK_PO_ID")
        
        budgetCheckResult.PO_ID = wePOId.Text
        
        PeopleSoft_PurchaseOrder_BudgetCheckResult_ExtractFromPage driver, budgetCheckResult
    End If


    PeopleSoft_PurchaseOrder_SaveWithBudgetCheck = True
    Exit Function
    ' ---------------------------------------------------------------------
    ' End - Save w/ Budget Check
    ' ---------------------------------------------------------------------
    
ExceptionThrown:
    PeopleSoft_PurchaseOrder_SaveWithBudgetCheck = False
    Err.Raise Err.Number, Err.Source, "PeopleSoft_PurchaseOrder_SaveWithBudgetCheck Exception: " & Err.Description, Err.Helpfile, Err.HelpContext
    

End Function
Public Function PeopleSoft_PurchaseOrder_BudgetCheckResult_ExtractFromPage(driver As SeleniumWrapper.WebDriver, ByRef budgetCheckResult As PeopleSoft_PurchaseOrder_BudgetCheckResult) As Boolean

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
    
    
    budgetCheckResult.BudgetCheck_HasErrors = True
    
    ' Begin - Line Errors
    PO_ErrorCount = CInt(driver.getXpathCount(".//*[contains(@id,'trZ_KK_PO_ERR_VW$0_row')]"))
    
    budgetCheckResult.BudgetCheck_Errors.BudgetCheck_LineErrorCount = PO_ErrorCount
    
    
    ReDim budgetCheckResult.BudgetCheck_Errors.BudgetCheck_LineErrors(1 To PO_ErrorCount) As PeopleSoft_PurchaseOrder_BudgetCheck_LineError

    
    For i = 1 To PO_ErrorCount
        PO_ErrorIndex = i - 1
        
        With budgetCheckResult.BudgetCheck_Errors.BudgetCheck_LineErrors(i)
            .LINE_NBR = CInt(driver.findElementById("Z_KK_PO_ERR_VW_LINE_NBR$" & PO_ErrorIndex).Text)
            .SCHED_NBR = CInt(driver.findElementById("Z_KK_PO_ERR_VW_SCHED_NBR$" & PO_ErrorIndex).Text)
            .DISTRIB_LINE_NUM = CInt(driver.findElementById("Z_KK_PO_ERR_VW_DISTRIB_LINE_NUM$" & PO_ErrorIndex).Text)
            .BUDGET_DT = driver.findElementById("Z_KK_PO_ERR_VW_BUDGET_DT$" & PO_ErrorIndex).Text
            .BUSINESS_UNIT_PC = driver.findElementById("Z_KK_PO_ERR_VW_BUSINESS_UNIT_PC$" & PO_ErrorIndex).Text
            .PROJECT_ID = driver.findElementById("Z_KK_PO_ERR_VW_PROJECT_ID$" & PO_ErrorIndex).Text
            .LINE_AMOUNT = CurrencyFromString(driver.findElementById("Z_KK_PO_ERR_VW_MONETARY_AMOUNT$" & PO_ErrorIndex).Text)
            .COMMIT_AMT = CurrencyFromString(driver.findElementById("Z_KK_ERR_WRK_Z_COMMIT_AMT$" & PO_ErrorIndex).Text)
            .NOT_COMMIT_AMT = CurrencyFromString(driver.findElementById("Z_KK_ERR_WRK_Z_NOT_COMMIT_AMT$" & PO_ErrorIndex).Text)
            .AVAIL_BUDGET_AMT = CurrencyFromString(driver.findElementById("Z_KK_PO_ERR_VW_Z_BUDGET_AMT$" & PO_ErrorIndex).Text)
        End With
    Next i
    ' End - Line Errors
    
    ' Begin - Project Errors
    PO_ErrorCount = CInt(driver.getXpathCount(".//*[contains(@id,'trZ_KK_PRJ_ERR_VW$0_row')]"))
    
    budgetCheckResult.BudgetCheck_Errors.BudgetCheck_ProjectErrorCount = PO_ErrorCount
    
    
    ReDim budgetCheckResult.BudgetCheck_Errors.BudgetCheck_ProjectErrors(1 To PO_ErrorCount) As PeopleSoft_PurchaseOrder_BudgetCheck_ProjectError

    
    ' Extract Project Budget Check Errors from field
    For i = 1 To PO_ErrorCount
        PO_ErrorIndex = i - 1
        
        With budgetCheckResult.BudgetCheck_Errors.BudgetCheck_ProjectErrors(i)
            .BUSINESS_UNIT_PC = driver.findElementById("Z_KK_PRJ_ERR_VW_BUSINESS_UNIT_PC$" & PO_ErrorIndex).Text
            .PROJECT_ID = driver.findElementById("Z_KK_PRJ_ERR_VW_PROJECT_ID$" & PO_ErrorIndex).Text
            .NOT_COMMIT_AMT = CurrencyFromString(driver.findElementById("Z_KK_ERR_WRK_Z_NOT_COMMIT_AMT2$" & PO_ErrorIndex).Text)
            .AVAIL_BUDGET_AMT = CurrencyFromString(driver.findElementById("Z_KK_PRJ_ERR_VW_Z_BUDGET_AMT$" & PO_ErrorIndex).Text)
            .FUNDING_NEEDED = CurrencyFromString(driver.findElementById("Z_KK_ERR_WRK_Z_KK_BAL_AMT$" & PO_ErrorIndex).Text)
        End With
    Next i
    ' End - Project Errors
    
    PeopleSoft_PurchaseOrder_BudgetCheckResult_ExtractFromPage = True
    Exit Function


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
Private Function PeopleSoft_Page_GetElementText(driver As SeleniumWrapper.WebDriver, ByVal elementID As String) As Variant

    elementID = Replace(elementID, "'", "\'")

    PeopleSoft_Page_GetElementText = driver.executeScript("return document.getElementById('" & elementID & "').textContent;")

End Function
Private Function PeopleSoft_Page_GetElementValue(driver As SeleniumWrapper.WebDriver, ByVal elementID As String) As Variant

    elementID = Replace(elementID, "'", "\'")

    PeopleSoft_Page_GetElementValue = driver.executeScript("return document.getElementById('" & elementID & "').value;")

End Function

Public Sub PeopleSoft_Page_WaitForProcessing(driver As SeleniumWrapper.WebDriver, Optional timeout_s As Long = 60)

    
    Const POLL_INTERVAL_MS As Double = 500 ' 0.5 s
    
    Dim iter As Integer
    Dim loader_inProcess As Boolean, proc_visibility As Variant
    
    Dim MAX_ITER As Double
    
    MAX_ITER = timeout_s * 1000 / POLL_INTERVAL_MS
    
    iter = 0
    
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
    Loop Until iter > MAX_ITER Or (proc_visibility <> "visible" And loader_inProcess = False)
    

    If iter > MAX_ITER Then
        Err.Raise 513, , "PeopleSoft_Page_WaitForProcessing Timeout"
    End If
    

End Sub
Public Function PeopleSoft_Page_CheckForModal(driver As SeleniumWrapper.WebDriver) As Integer
    ' Returns index # of modal if found (starts at 0). Returns -1 if not found, -2 if error
    
    
    On Error GoTo NotFoundOrErr
    
    Dim wePopupModals As WebElementCollection
    
    
    Set wePopupModals = driver.findElementsByXPath(".//*[starts-with(@id,'ptMod_')]", 100)
    
    If wePopupModals.Count = 0 Then
        PeopleSoft_Page_CheckForModal = -1
        Exit Function
    End If
    
    
    Dim elemID As String, modalIndexStr As String
    
    elemID = wePopupModals(0).getAttribute("id")
    modalIndexStr = Right$(elemID, Len(elemID) - Len("ptMod_"))
    
    If Not IsNumeric(modalIndexStr) Then
        PeopleSoft_Page_CheckForModal = -2
        Exit Function
    End If
    
    PeopleSoft_Page_CheckForModal = CLng(modalIndexStr)
        
    Exit Function
    
NotFoundOrErr:
    PeopleSoft_Page_CheckForModal = -2
    

End Function
Public Function PeopleSoft_Page_CheckForPopup(driver As SeleniumWrapper.WebDriver, Optional acknowledgePopup As Boolean = False, Optional raiseErrorIfUnexpected As Boolean = True, Optional expectedContent As Variant) As PeopleSoft_Page_PopupCheckResult

    If DEBUG_OPTIONS.CaptureExceptions Then On Error GoTo PopupNotFoundOrErr
    
    
    Dim popupCheckResult As PeopleSoft_Page_PopupCheckResult
    
    
    popupCheckResult.HasPopup = False


    Dim we As SeleniumWrapper.WebElement, By As New SeleniumWrapper.By
    Dim wePopupModals As WebElementCollection
    
    Dim t0 As Double
    t0 = Timer
    
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
    popupCheckResult.PopupText = we.Text
    
    
    
    ' Check to see if the popup text matches any the patterns in expectContent - allow for either array or string
    Dim expectedPatterns() As Variant, expectedPattern As Variant
    Dim expectedDebugStr As String, i As Long
    
    expectedDebugStr = "NULL"
     
    ' Compare popup text with any of the strings listed in expectedContent() to determine if popup is expected
    If Not IsMissing(expectedContent) Then
        If IsArray(expectedContent) Then
            expectedPatterns = expectedContent
        Else
            expectedPatterns = Array(expectedContent)
        End If
        
        expectedDebugStr = "'" & Join(expectedPatterns, "','" & "'")
        
        For Each expectedPattern In expectedPatterns
            If popupCheckResult.PopupText Like CStr(expectedPattern) Then
                popupCheckResult.IsExpected = True
                Exit For
            End If
        Next
    End If
        
    
    Debug.Print "PeopleSoft_Page_CheckForPopup: ID='" & popupCheckResult.PopupElementID & "', Expected=" & popupCheckResult.IsExpected & ", " _
                & "Buttons=(" & IIf(popupCheckResult.HasButtonYes, "Yes", "") & IIf(popupCheckResult.HasButtonNo, "|No", "") & IIf(popupCheckResult.HasButtonOk, "|OK", "") & IIf(popupCheckResult.HasButtonCancel, "|Cancel", "") & "), " _
                & "Text='" & popupCheckResult.PopupText & "', ExpectedContents=" & expectedDebugStr & ""
    
    If raiseErrorIfUnexpected And Not IsMissing(expectedContent) And popupCheckResult.IsExpected = False Then
        On Error GoTo 0
        Err.Raise -1, , "Unexpected Popup: " & popupCheckResult.PopupText
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
    popupCheckResult.PopupText = ""
    
    PeopleSoft_Page_CheckForPopup = popupCheckResult
    
    Debug.Print "PeopleSoft_Page_CheckForPopup: No popup found or error: " & Err.Description

End Function
Public Sub PeopleSoft_Page_AcknowledgePopup(driver As SeleniumWrapper.WebDriver, ByRef popupCheckResult As PeopleSoft_Page_PopupCheckResult, clickButton As VbMsgBoxResult)
    
    If DEBUG_OPTIONS.AddFunctionNameToExceptions Then On Error GoTo ExceptionThrown
    
    If clickButton = vbOK Then
        driver.findElementByXPath(".//*[@id='" & popupCheckResult.PopupElementID & "']/descendant::*[@id='#ICOK']").Click
    ElseIf clickButton = vbCancel Then
        driver.findElementByXPath(".//*[@id='" & popupCheckResult.PopupElementID & "']/descendant::*[@id='#ICCancel']").Click
    ElseIf clickButton = vbYes Then
        driver.findElementByXPath(".//*[@id='" & popupCheckResult.PopupElementID & "']/descendant::*[@id='#ICYes']").Click
    ElseIf clickButton = vbNo Then
        driver.findElementByXPath(".//*[@id='" & popupCheckResult.PopupElementID & "']/descendant::*[@id='#ICNo']").Click
    End If
    
    PeopleSoft_Page_WaitForProcessing driver
    
    Exit Sub
    
ExceptionThrown:
    Err.Raise Err.Number, Err.Source, "PeopleSoft_Page_AcknowledgePopup: " & Err.Description, Err.Helpfile, Err.HelpContext

End Sub
' PeopleSoft_Page_SuppressPopup: Wrapper function to acknowledge popup and return only the text. This is deprecated but retained for backward compatibility. Use PeopleSoft_Page_CheckForPopup instead
Public Function PeopleSoft_Page_SuppressPopup(driver As SeleniumWrapper.WebDriver, clickButton As VbMsgBoxResult, Optional matchText As String = "") As String

    Dim popupCheckResult As PeopleSoft_Page_PopupCheckResult


    If DEBUG_OPTIONS.AddFunctionNameToExceptions Then On Error GoTo ExceptionThrown
    

    popupCheckResult = PeopleSoft_Page_CheckForPopup(driver)
    If popupCheckResult.HasPopup = False Then Exit Function
    
    PeopleSoft_Page_SuppressPopup = popupCheckResult.PopupText
    
    If matchText <> "" Then
        If Not popupCheckResult.PopupText Like matchText Then
            Debug.Print "PeopleSoft_Page_SuppressPopup: Unexpected popup. Text does not match '" & matchText & "'"
            Err.Raise -1, , "PeopleSoft_Page_SuppressPopup: Unexpected popup. Text does not match." & vbCrLf & "Popup Text: " & popupCheckResult.PopupText & vbCrLf & "Match: " & matchText & ""
            Exit Function
        End If
    End If
    
    PeopleSoft_Page_AcknowledgePopup driver, popupCheckResult, clickButton
    
    Exit Function

ExceptionThrown:
    Err.Raise Err.Number, Err.Source, "PeopleSoft_Page_SuppressPopup: " & Err.Description, Err.Helpfile, Err.HelpContext

End Function


Private Function CurrencyFromString(strCur As String) As Currency

    strCur = Replace(strCur, ",", "")
    
    If IsNumeric(strCur) Then
        CurrencyFromString = CCur(strCur)
    Else
        CurrencyFromString = 0
    End If

End Function



