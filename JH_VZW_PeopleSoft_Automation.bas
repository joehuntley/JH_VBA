Attribute VB_Name = "JH_VZW_PeopleSoft_Automation"
Option Explicit

' PeopleSoft Automation Module
' used Selenium to automate PS UI
'
' joseph.huntley@vzw.com


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
    
    ' Field Validation Results
    PO_BUSINESS_UNIT_Result As PeopleSoft_Field_ValidationResult
    VENDOR_NAME_SHORT_Result As PeopleSoft_Field_ValidationResult
    PO_HDR_VENDOR_ID_Result As PeopleSoft_Field_ValidationResult
    PO_HDR_VENDOR_LOCATION_Result As PeopleSoft_Field_ValidationResult
    PO_HDR_BUYER_ID_Result As PeopleSoft_Field_ValidationResult
    PO_HDR_APPROVER_ID_Result As PeopleSoft_Field_ValidationResult
    
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
    
    ITEM_ID As String
    TRANS_ITEM_DESC As String
    
    RECEIVE_QTY As Variant ' Decimal type
    ACCEPT_QTY As Variant ' Decimal type
    
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
    ReceiptItemCount As Integer
    
    HasGlobalError As Boolean
    GlobalError As String
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

Public Function GetSeleniumVersion() As String

    Dim assy As New SeleniumWrapper.Assembly
    
    GetSeleniumVersion = assy.GetVersion
    

End Function



' -----------------------------------------------------------------------------------------------
Public Function PeopleSoft_NewSession(user As String, pass As String) As PeopleSoft_Session

    Dim session As PeopleSoft_Session
    Dim driver As New SeleniumWrapper.WebDriver
    
    
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
        driver.setImplicitWait 3000
        
        
        driver.get PS_URI_LOGIN
        
          
        driver.findElementByName("USER").Clear
        driver.findElementByName("USER").SendKeys session.user
        driver.findElementByName("password").Clear
        driver.findElementByName("password").SendKeys session.pass
        driver.findElementByName("btn_logon").Click
        
        driver.waitForPageToLoad 5000 ' wait up to 5s
        
        
        
        Dim By As New SeleniumWrapper.By, weErrorBoxMsg As SeleniumWrapper.WebElement
        Dim errMsg As String
        
        
        If PeopleSoft_Page_ElementExists(driver, By.XPath(".//*[@id='ErrorBox']/tbody/tr/td/font/b")) Then
            errMsg = driver.findElementByXPath(".//*[@id='ErrorBox']/tbody/tr/td/font/b").Text
                    
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

End Function
Public Sub PeopleSoft_NavigateTo_ExistingPO(ByRef session As PeopleSoft_Session, PO_BU As String, PO_ID As String)
    
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
    
End Sub
Public Function PeopleSoft_PurchaseOrder_CutPO(ByRef session As PeopleSoft_Session, ByRef purchaseOrder As PeopleSoft_PurchaseOrder) As Boolean

   ' On Error GoTo ExceptionThrown


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
    
    ' -------------------------------------------------------------------
    ' Begin - Comments Section
    ' -------------------------------------------------------------------
    If Len(purchaseOrder.PO_Fields.PO_HDR_COMMENTS) > 0 Then ' Only go into comments section if commend is given
        driver.findElementById("COMM_WRK1_COMMENTS_PB").Click
        'driver.runScript "javascript:hAction_win0(document.win0,'COMM_WRK1_COMMENTS1_PB', 0, 0, 'Edit Comments', false, true);"
        
        'driver.waitForPageToLoad 5000
        PeopleSoft_Page_WaitForProcessing driver
         
        If False Then ' No more suggested approver
            driver.waitForElementPresent "css=#PO_HDR_Z_SUG_APPRVR"
            
            
            PeopleSoft_Page_SetValidatedField driver, ("PO_HDR_Z_SUG_APPRVR"), _
                CStr(purchaseOrder.PO_Fields.PO_HDR_APPROVER_ID), purchaseOrder.PO_Fields.PO_HDR_APPROVER_ID_Result
            
            If purchaseOrder.PO_Fields.PO_HDR_APPROVER_ID_Result.ValidationFailed Then GoTo ValidationFail
        End If
        
        If Len(purchaseOrder.PO_Fields.PO_HDR_COMMENTS) > 0 Then
            driver.findElementById("COMM_WRK1_COMMENTS_2000$0").Clear
            driver.findElementById("COMM_WRK1_COMMENTS_2000$0").SendKeys purchaseOrder.PO_Fields.PO_HDR_COMMENTS
        End If
        
        PeopleSoft_Page_WaitForProcessing driver
        
    
        driver.findElementById("#ICSave").Click
        'driver.runScript "javascript:submitAction_win0(document.win0, '#ICSave');" ' work-around - Clicks 'Save'
        
        PeopleSoft_Page_WaitForProcessing driver
    
    End If
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
    'driver.runScript "javascript:hAction_win0(document.win0,'PO_PNLS_PB_EXPAND_ALL_PB', 0, 0, 'Expand All', false, true);"
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
        'driver.runScript "javascript:hAction_win0(document.win0,'PO_PNLS_PB_EXPAND_ALL_PB', 0, 0, 'Expand All', false, true);"
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

    On Error GoTo ExceptionThrown
    
    Output_Print "PeopleSoft_PurchaseOrder_CreateFromQuote"
    Output_Indent_Increase
    
    


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
    
    
    PeopleSoft_Page_SetValidatedField driver, ("PO_HDR_BUYER_ID"), _
        CStr(poCFQ.PO_Fields.PO_HDR_BUYER_ID), poCFQ.PO_Fields.PO_HDR_BUYER_ID_Result
        
    If poCFQ.PO_Fields.PO_HDR_BUYER_ID_Result.ValidationFailed Then GoTo ValidationFail
    
    
    If Len(poCFQ.PO_Fields.PO_HDR_PO_REF) > 0 Then
        driver.findElementById("PO_HDR_PO_REF").Clear
        driver.findElementById("PO_HDR_PO_REF").SendKeys poCFQ.PO_Fields.PO_HDR_PO_REF
    End If
    
    
    'Dim elemSelect As SeleniumWrapper.Select
    Dim elemSelect As SeleniumWrapper.WebElement
    Dim selectOptions As WebElementCollection, selectOptionsElement As WebElement
    
    Set elemSelect = driver.findElementById("PO_COPY_TMPLT_W_COPY_PO_FROM")
    
    Debug.Print "PO_COPY_TMPLT_W_COPY_PO_FROM - Options"
    
    Set selectOptions = elemSelect.AsSelect.Options
    
    For Each selectOptionsElement In selectOptions
        Debug.Print "- " & selectOptionsElement.getAttribute("value") & ":" & selectOptionsElement.Text
    Next selectOptionsElement
    
    
    'elemSelect.Click
    Dim tryNo As Long
    
    ' Select Copy Purchase Order from eQuote - Try a few different ways
    For tryNo = 1 To 5
        Debug.Print "Selecting from Dropdown: Try #" & tryNo
            
        Set elemSelect = driver.findElementById("PO_COPY_TMPLT_W_COPY_PO_FROM")
        
        Debug.Print elemSelect.AsSelect.Options
        
    
        Select Case tryNo
            Case 1:
                Debug.Print "- Method: AsSelect.selectByText"
                elemSelect.AsSelect.selectByText "eQuote"
            Case 2:
                Debug.Print "- Method: AsSelect.selectByValue"
                elemSelect.AsSelect.selectByValue "Q"
            Case 3:
                Debug.Print "- Method: JS: Set value"
                driver.runScript "javascript: document.getElementById('PO_COPY_TMPLT_W_COPY_PO_FROM').value = 'Q';"
            Case 4:
                Debug.Print "- Method: JS: submitAction"
                driver.runScript "javascript: var elem = document.getElementById('PO_COPY_TMPLT_W_COPY_PO_FROM'); addchg_win0(elem); submitAction_win0(elem.form,elem.name);"
            Case 5:
                Debug.Print "- Method: JS: invoke onChange"
                driver.runScript "javascript: var elem = document.getElementById('PO_COPY_TMPLT_W_COPY_PO_FROM'); elem.onchange();"
        End Select
        
        PeopleSoft_Page_WaitForProcessing driver
        
        If PeopleSoft_Page_ElementExists(driver, By.XPath(".//*[text()='Create from Quote']")) Then
            Exit For
        End If
    Next tryNo
    
    'Debug.Print
    'driver.runScript "javascript: var elem = document.getElementById('PO_COPY_TMPLT_W_COPY_PO_FROM'); addchg_win0(elem); submitAction_win0(elem.form,elem.name);"
    'Debug.Print
    'driver.runScript "javascript: var elem = document.getElementById('PO_COPY_TMPLT_W_COPY_PO_FROM'); elem.onchange();"
   
    ' <h1 class="PSSRCHTITLE">Create from Quote</h1>
    driver.waitForElementPresent "xpath=.//*[text()='Create from Quote']"
    

    ' Type Vendor ID
    PeopleSoft_Page_SetValidatedField driver, ("Z_E_QT_WRK_VENDOR_ID"), _
        Format(poCFQ.PO_Fields.PO_HDR_VENDOR_ID, "0000000000"), poCFQ.PO_Fields.PO_HDR_VENDOR_ID_Result
    
    If poCFQ.PO_Fields.PO_HDR_VENDOR_ID_Result.ValidationFailed Then GoTo ValidationFail
    
    ' Type Quote Number
    PeopleSoft_Page_SetValidatedField driver, ("Z_E_QT_WRK_Z_QUOTE_NBR"), _
        poCFQ.E_QUOTE_NBR, poCFQ.E_QUOTE_NBR_Result
    
    If poCFQ.E_QUOTE_NBR_Result.ValidationFailed Then GoTo ValidationFail
    
    ' Click Search
    driver.findElementById("Z_E_QT_WRK_REFRESH").Click
    'driver.runScript "javascript:hAction_win0(document.win0,'Z_E_QT_WRK_REFRESH', 0, 0, 'Search', false, true)"
    PeopleSoft_Page_WaitForProcessing driver
    
    
    Dim loadedQuoteNbr As String
    loadedQuoteNbr = driver.findElementById("Z_E_QT_CPPO_VW_Z_QUOTE_NBR$0").Text

    ' Sanity check
    If loadedQuoteNbr <> poCFQ.E_QUOTE_NBR Then
        poCFQ.HasError = True
        poCFQ.GlobalError = "Error loading quote: quote mismatch"
    End If
    
    
    ' Click Apply
    driver.findElementById("Z_E_QT_WRK_APPLY").Click
    'driver.runScript "javascript:hAction_win0(document.win0,'Z_E_QT_WRK_APPLY', 0, 0, 'Apply', false, true)"
    PeopleSoft_Page_WaitForProcessing driver, TIMEOUT_LONG


    PeopleSoft_PurchaseOrder_PO_Defaults_Fill driver, poCFQ.PO_Defaults
        
    If poCFQ.PO_Defaults.HasValidationError Then GoTo ValidationFail
    
    
    
    ' -------------------------------------------------------------------
    ' Begin - Comments Section
    ' -------------------------------------------------------------------
    If Len(poCFQ.PO_Fields.PO_HDR_COMMENTS) > 0 Then
        driver.findElementById("COMM_WRK1_COMMENTS_PB").Click
        'driver.runScript "javascript:hAction_win0(document.win0,'COMM_WRK1_COMMENTS1_PB', 0, 0, 'Edit Comments', false, true);"
        
        'driver.waitForPageToLoad 5000
        PeopleSoft_Page_WaitForProcessing driver
         
        If False Then ' No more PO approver.
            driver.waitForElementPresent "css=#PO_HDR_Z_SUG_APPRVR"
            
            
            PeopleSoft_Page_SetValidatedField driver, ("PO_HDR_Z_SUG_APPRVR"), _
                CStr(poCFQ.PO_Fields.PO_HDR_APPROVER_ID), poCFQ.PO_Fields.PO_HDR_APPROVER_ID_Result
            
            If poCFQ.PO_Fields.PO_HDR_APPROVER_ID_Result.ValidationFailed Then GoTo ValidationFail
        End If
        
        If Len(poCFQ.PO_Fields.PO_HDR_COMMENTS) > 0 Then
            driver.findElementById("COMM_WRK1_COMMENTS_2000$0").Clear
            driver.findElementById("COMM_WRK1_COMMENTS_2000$0").SendKeys poCFQ.PO_Fields.PO_HDR_COMMENTS
        End If
        
        PeopleSoft_Page_WaitForProcessing driver
        
    
        driver.findElementById("#ICSave").Click
        'driver.runScript "javascript:submitAction_win0(document.win0, '#ICSave');" ' work-around - Clicks 'Save'
        
        PeopleSoft_Page_WaitForProcessing driver
    
    End If
    ' -------------------------------------------------------------------
    ' End - Comments Section
    ' -------------------------------------------------------------------
    
    

    
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
    'driver.runScript "javascript:hAction_win0(document.win0,'CALCULATE_TAXES', 0, 0, 'Calculate', false, true);"
    
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
    PeopleSoft_PurchaseOrder_CreateFromQuote = False
    Exit Function
    
ExceptionThrown:
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


    Dim isAnyDefaultSpecified As Boolean
    
    Dim PopupText As String
    
    isAnyDefaultSpecified = False
    
    If PO_Defaults.SCH_DUE_DATE > 0 Then isAnyDefaultSpecified = True
    If PO_Defaults.SCH_SHIPTO_ID > 0 Then isAnyDefaultSpecified = True
    If Len(PO_Defaults.DIST_BUSINESS_UNIT_PC) > 0 Then isAnyDefaultSpecified = True
    If Len(PO_Defaults.DIST_PROJECT_CODE) > 0 Then isAnyDefaultSpecified = True
    If Len(PO_Defaults.DIST_ACTIVITY_ID) > 0 Then isAnyDefaultSpecified = True
    If PO_Defaults.DIST_LOCATION_ID > 0 Then isAnyDefaultSpecified = True

    
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
            PeopleSoft_Page_SetValidatedField driver, _
                ("PO_DFLT_TBL_DUE_DT"), _
                Format(PO_Defaults.SCH_DUE_DATE, "mm/dd/yyyy"), _
                PO_Defaults.SCH_DUE_DATE_Result
            
            If PO_Defaults.SCH_DUE_DATE_Result.ValidationFailed Then GoTo ValidationFail
        End If
    

        
        If Len(PO_Defaults.DIST_BUSINESS_UNIT_PC) > 0 Then
            PeopleSoft_Page_SetValidatedField driver, _
                ("BUSINESS_UNIT_PC$0"), _
                PO_Defaults.DIST_BUSINESS_UNIT_PC, _
                PO_Defaults.DIST_BUSINESS_UNIT_PC_Result
            
            If PO_Defaults.DIST_BUSINESS_UNIT_PC_Result.ValidationFailed Then GoTo ValidationFail
        End If
        
        If Len(PO_Defaults.DIST_PROJECT_CODE) > 0 Then
            PeopleSoft_Page_SetValidatedField driver, _
                ("PROJECT_ID$0"), _
                PO_Defaults.DIST_PROJECT_CODE, _
                PO_Defaults.DIST_PROJECT_CODE_Result
            
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
            PeopleSoft_Page_SetValidatedField driver, _
                ("PO_DFLT_DISTRIB_SHIPTO_ID$0"), _
                CStr(PO_Defaults.SCH_SHIPTO_ID), _
                PO_Defaults.SCH_SHIPTO_ID_Result
            
            If PO_Defaults.SCH_SHIPTO_ID_Result.ValidationFailed Then GoTo ValidationFail
        End If
        
        If PO_Defaults.DIST_LOCATION_ID > 0 Then
            PeopleSoft_Page_SetValidatedField driver, _
                ("LOCATION$0"), _
                CStr(PO_Defaults.DIST_LOCATION_ID), _
                PO_Defaults.DIST_LOCATION_ID_Result
            
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
    
    
    

End Function
Private Function PeopleSoft_PurchaseOrder_SetValidatedField_ActivityID(driver As SeleniumWrapper.WebDriver, activityID_elementID As String, activityID_value As String, ByRef activityID_validationResult As PeopleSoft_Field_ValidationResult, activityIdPrompt_ElementID As String) As String

    'On Error GoTo ErrOccurred
    
    
    Dim activityListString As String: activityListString = ""

    PeopleSoft_Page_SetValidatedField driver, activityID_elementID, activityID_value, activityID_validationResult

    Exit Function ' DO NOT CONTINUE (until below code is fixed)
    
    ' TODO: FIX THE CODE BELOW
    If activityID_validationResult.ValidationFailed Then
        If InStr(1, activityID_validationResult.ValidationErrorText, "Invalid value") > 0 Then
            Dim tmpFVR As PeopleSoft_Field_ValidationResult
            
            PeopleSoft_Page_SetValidatedField driver, activityID_elementID, "", tmpFVR, False
        
            'activityListString = PeopleSoft_PurchaseOrder_Extract_ActivityIDs(driver, activityIdPrompt_ElementID)

            'Dim activityListArr() As String
            Dim activityListCount As Integer
            
            ' Simulates clicking on the spyglass. Extracts the activity IDs from the popup.
            driver.findElementById(activityIdPrompt_ElementID).Click
            'driver.runScript "javascript:pAction_win0(document.win0," & activityIdPrompt_ElementID & ");"
            PeopleSoft_Page_WaitForProcessing driver
            
            'driver.waitForElementPresent "css=#popupFrame"
            'driver.waitForElementPresent "css=#PTPOPUP_TITLE"
            'driver.waitForElementPresent "css=.PTPOPUP_TITLE"
            driver.waitForElementPresent "css=#ptModTitle_1"
            
            
            'Dim weFrame As WebElement
            
            'Set weFrame = driver.findElementById("popupFrame")
            
            'driver.switchToFrame "#popupFrame"
            'driver.switchToFrame "#ptMod_1"
            
            
        
            Dim webElemsActivityIDs As SeleniumWrapper.WebElementCollection
            Dim webElemsActivityDescriptions As SeleniumWrapper.WebElementCollection
            
            'Set webElemsActivityIDs = driver.findElementsByXPath(".//*[@id='win0divSEARCHRESULT']/descendant::table[@class='PSSRCHRESULTSWBO']/tbody/tr/td[3]/a")
            'Set webElemsActivityDescriptions = driver.findElementsByXPath(".//*[@id='win0divSEARCHRESULT']/descendant::table[@class='PSSRCHRESULTSWBO']/tbody/tr/td[4]/a")
            
            Set webElemsActivityIDs = driver.findElementsByXPath(".//*[@id='win0divSEARCHRESULT']/descendant::table[@class='PSSRCHRESULTSWBO']/descendant::a[contains(@class,'RESULT4$')]")
            Set webElemsActivityDescriptions = driver.findElementsByXPath(".//*[@id='win0divSEARCHRESULT']/descendant::table[@class='PSSRCHRESULTSWBO']/descendant::a[contains(@class,'RESULT5$')]")
           
            
            activityListCount = webElemsActivityIDs.Count
            
            If activityListCount > 0 Then
                ReDim activityListArr(1 To activityListCount) As String
                
                Dim i As Integer
                
                For i = 1 To activityListCount
                    activityListString = activityListString & webElemsActivityIDs(i - 1).Text & " (" & webElemsActivityDescriptions(i - 1).Text & ")" & ","
                Next i
            End If
            
            If Len(activityListString) > 0 Then
                activityListString = Left(activityListString, Len(activityListString) - 1)
                activityID_validationResult.ValidationErrorText = "Invalid activity ID. Valid values are as follows: " & activityListString
            End If
        
            ' Close dialog
            driver.runScript "javascript:var ptc_tmp = new PT_common(); ptc_tmp.updatePrompt(document.win0,'#ICCancel');"
            
        
         'NOTES:
         '  Update Parent: <a name="RESULT4$1" id="RESULT4$1" href="javascript:doUpdateParent(document.win0,'#ICRow1');">CNSTR</a>
            'driver.selectFrame "relative=top"
        
            
           
    
    
        End If
    End If
    
    PeopleSoft_PurchaseOrder_SetValidatedField_ActivityID = activityListString
    
    Exit Function
    
ErrOccurred:

    activityID_validationResult.ValidationErrorText = activityID_validationResult.ValidationErrorText & vbCrLf & vbCrLf & "Exception: " & Err.Description
    
    

End Function
Public Function PeopleSoft_PurchaseOrder_Fill_PO_Line(driver As SeleniumWrapper.WebDriver, ByRef purchaseOrder As PeopleSoft_PurchaseOrder, PO_Line As Integer, ByVal PO_pageScheduleIndex As Integer) As Boolean

    Debug.Assert PO_Line > 0 And PO_Line <= purchaseOrder.PO_LineCount
    
    
        Debug.Print
    
    'On Error GoTo ExceptionThrown
    
    
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
            PeopleSoft_Page_SetValidatedField driver, ("PO_LINE_SHIP_DUE_DT$" & (PO_pageScheduleIndex + PO_Line_Schedule - 1)), _
                Format(purchaseOrder.PO_Lines(PO_Line).Schedules(PO_Line_Schedule).ScheduleFields.DUE_DATE, "mm/dd/yyyy"), _
                purchaseOrder.PO_Lines(PO_Line).Schedules(PO_Line_Schedule).ScheduleFields.DUE_DATE_Result
            
            
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
    purchaseOrder.GlobalError = "Exception: " & Err.Description
    
    PeopleSoft_PurchaseOrder_Fill_PO_Line = False
    

End Function


Public Function PeopleSoft_PurchaseOrder_ProcessChangeOrder(ByRef session As PeopleSoft_Session, ByRef poChangeOrder As PeopleSoft_PurchaseOrder_ChangeOrder) As Boolean
    
    
    'On Error GoTo ExceptionThrown
    
    
    Dim By As New By, Assert As New Assert, Verify As New Verify
    Dim driver As New SeleniumWrapper.WebDriver
    
    
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
        
        
        If PeopleSoft_Page_ElementExists(driver, By.ID("COMM_WRK1_COMMENTS1_PB")) Then
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
        
        
        ' Deprecated after 2.9.1.2
        'If PeopleSoft_Page_ElementExists(driver, By.LinkText("Add Comments")) Then
        '    Set weCmtsLink = driver.findElementByLinkText("Add Comments")
        'Else
        '    Set weCmtsLink = driver.findElementByLinkText("Edit Comments")
        'End If
        
        'elID = weCmtsLink.getAttribute("id")
        
        'driver.executeScript "javascript:arguments[0].focus();", weCmtsLink
        'weCmtsLink.Click
        

         
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
            
            
            If PeopleSoft_Page_ElementExists(driver, By.ID("PO_SCR_NAV_WRK_SRCH_RSLT_MSG")) Then
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
    If PeopleSoft_Page_ElementExists(driver, By.ID("PO_CHNG_REASON_COMMENTS$0")) Then
        
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


Public Function PeopleSoft_PurchaseOrder_ProcessReceipt(ByRef session As PeopleSoft_Session, ByRef rcpt As PeopleSoft_Receipt) As Boolean





    On Error GoTo ExceptionThrown

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
    
    driver.findElementById("PO_PICK_ORD_WRK_ORDER_ID").Clear
    driver.findElementById("PO_PICK_ORD_WRK_ORDER_ID").SendKeys rcpt.PO_ID
    
    
    driver.findElementById("PO_PICK_ORD_WRK_PB_FETCH_PO").Click
    
    
    PeopleSoft_Page_WaitForProcessing driver
    
    ' ------------------------------------------------------
    ' Begin - Map receivable  items on page to receipt items
    ' ------------------------------------------------------
    Dim rowIndexMap() As Integer
    
    ' update 2.10.1: pre-allocate index map array here for receiving on specific items
    If rcpt.ReceiveMode = RECEIVE_SPECIFIED Then
        ReDim rowIndexMap(1 To rcpt.ReceiptItemCount) As Integer
        
        For i = 1 To rcpt.ReceiptItemCount
            rowIndexMap(i) = -1
        Next i
    End If
    

    'If Not PeopleSoft_Page_ElementExists(driver, By.ID("win0divGPPO_PICK_ORD_WS$0")) Then
    If Not PeopleSoft_Page_ElementExists(driver, By.ID("win0divPO_PICK_ORD_WS$0")) Then ' fix 2.9.1.3
        rcpt.HasGlobalError = True
        rcpt.GlobalError = "No receivable items on this PO"
    
        GoTo ReceiptFailed
    End If
    
    
    If Not PeopleSoft_Page_ElementExists(driver, By.ID("PO_PICK_ORD_WRK_Z_IN_CATS_FLAG$0")) Then
        rcpt.HasGlobalError = True
        rcpt.GlobalError = "No receivable items on this PO"
        
        GoTo ReceiptFailed
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
            rcpt.HasGlobalError = True
            rcpt.GlobalError = "No receivable items on this PO"
            
            GoTo ReceiptFailed
        End If
    End If
    
    ' new in version 2.10.1:
    ' If we are receiving all lines, then we will return all line item information, rather than match the specific line items
    If rcpt.ReceiveMode = RECEIVE_ALL Then
        rcpt.ReceiptItemCount = numReturnedRows
        ReDim rcpt.ReceiptItems(1 To rcpt.ReceiptItemCount) As PeopleSoft_Receipt_Item
        ReDim rowIndexMap(1 To rcpt.ReceiptItemCount) As Integer
    End If
    
    
    Debug.Print "# of rows: " & numReturnedRows
    
    For rowIndex = 0 To numReturnedRows - 1
        Dim Row_PO_ID As String, Row_PO_Line As Long, Row_PO_Sch As Long, Row_PO_Qty As Long
        Dim Row_PO_ITEM_ID As String, Row_CATS_FLAG As String, Row_CheckDisabled As String
      
        ' workaround because driver.findElementById(X).Text doesn't always return a value and is very slow
        Row_PO_ID = driver.executeScript("return document.getElementById('PO_PICK_ORD_WS_PO_ID$" & rowIndex & "').textContent;")                'driver.findElementById("PO_PICK_ORD_WS_PO_ID$" & rowIndex).Text
        Row_PO_Line = CLng(driver.executeScript("return document.getElementById('PO_PICK_ORD_WS_LINE_NBR$" & rowIndex & "').textContent;"))     'CInt(driver.findElementById("PO_PICK_ORD_WS_LINE_NBR$" & rowIndex).Text)
        Row_PO_Sch = CLng(driver.executeScript("return document.getElementById('PO_PICK_ORD_WS_SCHED_NBR$" & rowIndex & "').textContent;"))     'CInt(driver.findElementById("PO_PICK_ORD_WS_SCHED_NBR$" & rowIndex).Text)
        Row_PO_Qty = CLng(driver.executeScript("return document.getElementById('QTY_PO$" & rowIndex & "').textContent;"))                       'CInt(driver.findElementById("QTY_PO$" & rowIndex).Text)
        Row_PO_ITEM_ID = driver.executeScript("return document.getElementById('PO_PICK_ORD_WS_INV_ITEM_ID$" & rowIndex & "').textContent;")        'driver.findElementById("PO_PICK_ORD_WS_INV_ITEM_ID$" & rowIndex).Text
        Row_CATS_FLAG = driver.executeScript("return document.getElementById('PO_PICK_ORD_WRK_Z_IN_CATS_FLAG$" & rowIndex & "').textContent;")                'driver.findElementById("PO_PICK_ORD_WS_PO_ID$" & rowIndex).Text
       
        Row_CheckDisabled = driver.executeScript("return document.getElementById('RECV_PO_SCHEDULE$" & rowIndex & "').disabled;")
       
        ' Slow and deprecated
       ' Row_PO_ID = driver.findElementById("PO_PICK_ORD_WS_PO_ID$" & rowIndex).Text
        'Row_PO_Line = CLng(driver.findElementById("PO_PICK_ORD_WS_LINE_NBR$" & rowIndex).Text)
        'Row_PO_Sch = CLng(driver.findElementById("PO_PICK_ORD_WS_SCHED_NBR$" & rowIndex).Text)
        'Row_PO_Qty = CLng(driver.findElementById("QTY_PO$" & rowIndex).Text)
        'Row_PO_Item = driver.findElementById("PO_PICK_ORD_WS_INV_ITEM_ID$" & rowIndex).Text
        'Debug.Print Row_PO_ID & vbTab & Format(Row_PO_Line, "00") & "." & Format(Row_PO_Sch, "00") & vbTab & Format(Row_PO_Qty, "#0") & vbTab & Row_PO_Item
        
        If rcpt.ReceiveMode = RECEIVE_ALL Then
            ' new in version 2.10.0 (receive all lines): copy line information to receiptLines, rather than map
            j = rowIndex + 1
            rowIndexMap(j) = rowIndex
            
            rcpt.ReceiptItems(j).PO_Line = Row_PO_Line
            rcpt.ReceiptItems(j).PO_Schedule = Row_PO_Sch
            rcpt.ReceiptItems(j).RECEIVE_QTY = Row_PO_Qty
            rcpt.ReceiptItems(j).ACCEPT_QTY = Row_PO_Qty
            rcpt.ReceiptItems(j).ITEM_ID = Row_PO_ITEM_ID
            rcpt.ReceiptItems(j).CATS_FLAG = Row_CATS_FLAG
        Else
         ' receive specified: map each row to the corresponding specific line/schedule in ReceiptItems()
            For j = 1 To rcpt.ReceiptItemCount
                If rcpt.ReceiptItems(j).PO_Line = Row_PO_Line And rcpt.ReceiptItems(j).PO_Schedule = Row_PO_Sch Then
                    'Debug.Assert rowIndexMap(j) = -1
                    
                    ' If ITEM ID is specified, check to make sure the ITEM ID matches as well
                    If rcpt.ReceiptItems(j).ITEM_ID = "" Or rcpt.ReceiptItems(j).ITEM_ID = Row_PO_ITEM_ID Then
                        rowIndexMap(j) = rowIndex
                        Exit For
                    End If
                    
                End If
            Next j
        End If
        
    Next rowIndex
    ' ------------------------------------------------------
    ' End - Map receivable  items on page to receipt items
    ' ------------------------------------------------------
    
    Debug.Print
    
    Dim numUnmatchedItems As Integer: numUnmatchedItems = 0
    
    
    ' Go through mapping/receive items. Click checkbox to receive.
    ' Check if any of the receipt items have not been mapped. If so,
    ' it has already been received or it is not receivable by the user
    
    
    For i = 1 To rcpt.ReceiptItemCount
        If rowIndexMap(i) >= 0 Then
            ' HACK - Does not work for some reason. Element has to be "clicked" or the form values does not save
            'driver.runScript "javascript: document.getElementById('RECV_PO_SCHEDULE$" & rowIndexMap(i) & "').checked = true;"
            'driver.runScript "javascript: setupTimeout2(); var elem = document.getElementById('RECV_PO_SCHEDULE$" & rowIndexMap(i) & "'); " & _
                                        "document.getElementById('RECV_PO_SCHEDULE$chk$" & rowIndexMap(i) & "').value = 'Y'; " & _
                                        "doFocus_win0(elem,false,true);"
                                        
            ' Gather information
            If rcpt.ReceiptItems(i).ITEM_ID = "" Then rcpt.ReceiptItems(i).ITEM_ID = driver.findElementById("PO_PICK_ORD_WS_INV_ITEM_ID$" & rowIndexMap(i)).Text
            rcpt.ReceiptItems(i).TRANS_ITEM_DESC = driver.findElementById("PO_PICK_ORD_WS_DESCR254_MIXED$" & rowIndexMap(i)).Text
                              
                              
            ' Check the box
            Set elem = driver.findElementById("RECV_PO_SCHEDULE$" & rowIndexMap(i))
            
            
            If elem.getAttribute("disabled") <> "disabled" Then
                elem.Click '- Note: does not work if element not visible
            Else
                rcpt.ReceiptItems(i).IsNotReceivable = True
                rcpt.ReceiptItems(i).RECEIVE_QTY = 0
            End If
            
            rcpt.ReceiptItems(i).HasError = False
        Else
            rcpt.ReceiptItems(i).HasError = True
            rcpt.ReceiptItems(i).ItemError = "Cannot receive on this item: not receivable or already received."
            
            numUnmatchedItems = numUnmatchedItems + 1
        End If
    Next i
    
    If numUnmatchedItems = rcpt.ReceiptItemCount Then
        rcpt.HasGlobalError = True
        rcpt.GlobalError = "No items can be received on PO."
        
        GoTo ReceiptFailed
    End If

    
    'driver.findElementById("#ICSave").Click
    driver.runScript "javascript:submitAction_win0(document.win0, '#ICSave');"


    PeopleSoft_Page_WaitForProcessing driver
    
    


    ' Simulate "View All"
    driver.runScript "javascript:submitAction_win0(document.win0,'RECV_LN_SHIP$hviewall$0');"
    PeopleSoft_Page_WaitForProcessing driver
    
   
    
    
    Dim numRcptLines As Integer, rcptLineIndex As Integer
    
    numRcptLines = driver.getXpathCount(".//*[contains(@id,'ftrRECV_LN_SHIP$0_row')]")
    
    ' 2.10.1: updated to only check received count if specific lines are to be received.
    If rcpt.ReceiveMode = RECEIVE_SPECIFIED And numRcptLines <> rcpt.ReceiptItemCount - numUnmatchedItems Then
        rcpt.HasGlobalError = True
        rcpt.GlobalError = "Partially received: number of receipt lines does not match"
        
        GoTo ReceiptFailed
    End If
    
    
    ' Note: On next page the receipt items are displayed in order of Line, Schedule
    ' ...OR if not, we check item IDs and PO lines to ensure they match
    ' -----------------------------------------------------------
    ' Begin - Sort Receipt Items by  Line,Schedule (Bubble Sort Algorithm)
    ' -----------------------------------------------------------
    Dim rcptItemsSortedIdx() As Integer
    ReDim rcptItemsSortedIdx(1 To rcpt.ReceiptItemCount) As Integer
    
    
    
    ' Sort by Line,Schedule (Bubble Sort Algorithm)
    For i = 1 To rcpt.ReceiptItemCount
        rcptItemsSortedIdx(i) = i
        Debug.Print i & ": " & rcpt.ReceiptItems(i).PO_Line
    Next i
    
    For i = 1 To rcpt.ReceiptItemCount
        Dim swapIdx As Integer, tmpIdx As Integer
        
        swapIdx = i
        
        For j = i + 1 To rcpt.ReceiptItemCount
            If rcpt.ReceiptItems(rcptItemsSortedIdx(j)).PO_Line < rcpt.ReceiptItems(rcptItemsSortedIdx(swapIdx)).PO_Line Then
                swapIdx = j
            ElseIf rcpt.ReceiptItems(rcptItemsSortedIdx(j)).PO_Line = rcpt.ReceiptItems(rcptItemsSortedIdx(swapIdx)).PO_Line Then
                If rcpt.ReceiptItems(rcptItemsSortedIdx(j)).PO_Schedule < rcpt.ReceiptItems(rcptItemsSortedIdx(swapIdx)).PO_Schedule Then
                    swapIdx = j
                End If
            End If
        Next j
        
        If swapIdx <> i Then
            tmpIdx = rcptItemsSortedIdx(i)
            rcptItemsSortedIdx(i) = rcptItemsSortedIdx(swapIdx)
            rcptItemsSortedIdx(swapIdx) = tmpIdx
        End If
    Next i
    ' -----------------------------------------------------------
    ' End - Sort Receipt Items by  Line,Schedule (Bubble Sort Algorithm)
    ' -----------------------------------------------------------
    
    ' -----------------------------------------------------------
    ' Begin - Sanity Check: if receipt lines match the input ReceiptLines. Adjust ReceiptQty and return AcceptQty as needed.
    ' -----------------------------------------------------------
    Dim rcptIdx As Integer, rcptLinePageIndex As Integer
    Dim anyItemHasErrors As Boolean
    
    Dim rcptLinePage_ITEM_ID As String, rcptLinePage_TRANS_ITEM_DESC As String
    Dim rcptLinePage_ACCEPT_QTY As Variant
    
    
    If rcpt.ReceiveMode = RECEIVE_SPECIFIED Then
        rcptLinePageIndex = 0
        anyItemHasErrors = False
        
        For i = 1 To rcpt.ReceiptItemCount
            rcptIdx = rcptItemsSortedIdx(i)
            
            If rcpt.ReceiptItems(rcptIdx).HasError = False And rcpt.ReceiptItems(rcptIdx).IsNotReceivable = False Then
            
                rcptLinePage_ITEM_ID = driver.executeScript("return document.getElementById('INV_ITEM_ID$" & rcptLinePageIndex & "').textContent;") 'driver.findElementById("INV_ITEM_ID$" & rcptLinePageIndex).Text
                rcptLinePage_TRANS_ITEM_DESC = driver.executeScript("return document.getElementById('DESCR$" & rcptLinePageIndex & "').textContent;") 'driver.findElementById("DESCR$" & rcptLinePageIndex).Text
                rcptLinePage_ACCEPT_QTY = CDec(driver.executeScript("return document.getElementById('RECV_LN_SHIP_QTY_SH_ACCPT$" & rcptLinePageIndex & "').textContent;")) 'CDec(driver.findElementById("RECV_LN_SHIP_QTY_SH_ACCPT$" & rcptLinePageIndex).Text)
                
                
                ' update 2.10.1: only perform sanity checks if receiving mode is for specific lines/schedules
                If rcpt.ReceiveMode = RECEIVE_SPECIFIED Then
                    ' First sanity check (match up Item IDs and Trans Item Desc)
                    If rcptLinePage_ITEM_ID <> rcpt.ReceiptItems(rcptIdx).ITEM_ID Then
                        rcpt.ReceiptItems(rcptIdx).HasError = True
                        rcpt.ReceiptItems(rcptIdx).ItemError = "Receipt line mismatch. ITEM ID: " & rcptLinePage_ITEM_ID & " (Expected: " & rcpt.ReceiptItems(rcptIdx).ITEM_ID & ")"
                        anyItemHasErrors = True
                    End If
                    
                    ' Partial match on Trans Item Desc - Disabled (causes issues)
                    'Dim transItemDescPartial As String
                    'transItemDescPartial = Replace(rcpt.ReceiptItems(rcptIdx).TRANS_ITEM_DESC, "  ", " ") ' Convert double spaces to single spaces.
                    'transItemDescPartial = Left(transItemDescPartial, Len(rcptLinePage_TRANS_ITEM_DESC))
                    
                    'If rcpt.ReceiptItems(rcptIdx).HasError = False And rcptLinePage_TRANS_ITEM_DESC <> transItemDescPartial Then
                        'rcpt.ReceiptItems(rcptIdx).HasError = True
                        'rcpt.ReceiptItems(rcptIdx).ItemError = "Receipt line mismatch. TRANS_ITEM_DESC: " & rcptLinePage_TRANS_ITEM_DESC & " (Expected: " & transItemDescPartial & ")"
                        'anyItemHasErrors = True
                    'End If
                
                    ' First sanity check passed
                    If rcpt.ReceiptItems(rcptIdx).HasError = False Then
                        rcpt.ReceiptItems(rcptIdx).ACCEPT_QTY = rcptLinePage_ACCEPT_QTY
                        
                        ' Second check: receive quantity is less than accept qty
                        If rcpt.ReceiptItems(rcptIdx).RECEIVE_QTY > 0 Then ' Receipt qty specified
                            If rcpt.ReceiptItems(rcptIdx).RECEIVE_QTY > rcpt.ReceiptItems(rcptIdx).ACCEPT_QTY Then
                                rcpt.ReceiptItems(rcptIdx).HasError = True
                                rcpt.ReceiptItems(rcptIdx).ItemError = "Receive qty is greater than accepted qty (Accept Qty: " & rcpt.ReceiptItems(rcptIdx).ACCEPT_QTY & ")"
                                anyItemHasErrors = True
                            End If
                        End If
                        
                        ' Second check passed
                        If rcpt.ReceiptItems(rcptIdx).HasError = False Then
                            If rcpt.ReceiptItems(rcptIdx).RECEIVE_QTY > 0 Then ' Receipt qty specified - otherwise receive all
                                Dim tmpValidationResult As PeopleSoft_Field_ValidationResult
                                
                                PeopleSoft_Page_SetValidatedField driver, ("RECV_LN_SHIP_QTY_SH_RECVD$" & rcptLinePageIndex), CStr(rcpt.ReceiptItems(rcptIdx).RECEIVE_QTY), tmpValidationResult
                                
                                If tmpValidationResult.ValidationFailed Then
                                    rcpt.ReceiptItems(rcptIdx).HasError = True
                                    rcpt.ReceiptItems(rcptIdx).ItemError = "RECEIVE QTY ERROR: " & tmpValidationResult.ValidationErrorText
                                    anyItemHasErrors = True
                                End If
                            Else
                                ' No receipt qty give. Receive on all and return the qty.
                                rcpt.ReceiptItems(rcptIdx).RECEIVE_QTY = CDec(driver.findElementById(("RECV_LN_SHIP_QTY_SH_RECVD$" & rcptLinePageIndex)).getAttribute("value"))
                            End If
                        End If 'Second check passed
                        
                    End If ' First sanity check passed
                    
                End If ' rcpt.ReceiveMode = RECEIVE_SPECIFIED
                
                
                rcptLinePageIndex = rcptLinePageIndex + 1
            End If
        Next i
        
        
        Dim PopupText As String, popUpIsExpected As Boolean
        
        
        
        If anyItemHasErrors Then
            rcpt.HasGlobalError = True
            rcpt.GlobalError = "Receipt lines matching error."
            
            
            ' Begin - Cancel Receipt
            driver.findElementById("RECV_HDR_WK_PB_CANCEL_RECPT").Click
            'driver.runScript "javascript:hAction_win0(document.win0,'RECV_HDR_WK_PB_CANCEL_RECPT', 0, 0, 'Cancel Receipt', false, true);"
            PeopleSoft_Page_WaitForProcessing driver
            
            
            PopupText = PeopleSoft_Page_SuppressPopup(driver, vbYes)
            'popUpIsExpected = InStr(1, popUpText, "Canceling Receipt cannot be reversed.") > 0
            
            'If popUpIsExpected = False Then
            '    rcpt.HasGlobalError = True
            '    rcpt.GlobalError = "Unexpected popup: " & popUpText
            '
            '    GoTo ReceiptFailed
            'End If
            
            PeopleSoft_Page_WaitForProcessing driver
            ' End - Cancel Receipt
            
            
            GoTo ReceiptFailed
        End If
    End If
    
    
    
    'driver.findElementById("#ICSave").Click
    driver.runScript "javascript:setSaveText_win0('Saving...');submitAction_win0(document.win0, '#ICSave');"


    ' Wait for "Saving..." to stop.
    driver.waitForElementPresent "css=#SAVED_win0"
    'driver.findElementById("processing").waitForCssValue "visibility", "visible"
    driver.findElementById("SAVED_win0").waitForCssValue "visibility", "hidden"

    
    
    Dim popupCheckResult As PeopleSoft_Page_PopupCheckResult

    Debug.Print "Expecting Popup: Have these receipt quantities been checked for accuracy"
    popupCheckResult = PeopleSoft_Page_CheckForPopup(driver)
    
    If popupCheckResult.HasPopup = False Or InStr(1, popupCheckResult.PopupText, "Have these receipt quantities been checked for accuracy") = 0 Then
        rcpt.HasGlobalError = True
        rcpt.GlobalError = "Did not receive expected popup: Have these receipt quantities been checked for accuracy?" _
                            & IIf(popupCheckResult.HasPopup, vbCrLf & "Popup received: " & popupCheckResult.PopupText, "")
        
        GoTo ReceiptFailed
    End If
    
    ' We received correct popup -> acknowledge
    PeopleSoft_Page_AcknowledgePopup driver, popupCheckResult, vbYes
    PeopleSoft_Page_WaitForProcessing driver
    
    'PopupText = PeopleSoft_Page_SuppressPopup(driver, vbYes)
    'popUpIsExpected = InStr(1, PopupText, "Have these receipt quantities been checked for accuracy") > 0
    
    'If popUpIsExpected = False Then
    '    rcpt.HasGlobalError = True
    '    rcpt.GlobalError = "Unexpected popup: " & PopupText
    '
    '    GoTo ReceiptFailed
    'End If
    
    
    ' Check for receipt ID.
    rcpt.RECEIPT_ID = driver.findElementById("RECV_HDR_RECEIVER_ID").Text
    rcpt.RECEIPT_ID = Trim(rcpt.RECEIPT_ID)
    Debug.Print "Receipt ID: " & rcpt.RECEIPT_ID
    
    
    
    If Not IsNumeric(rcpt.RECEIPT_ID) Then
        rcpt.HasGlobalError = True
        rcpt.GlobalError = "Non-numeric receipt ID not found on page: " & rcpt.RECEIPT_ID
    
        GoTo ReceiptFailed
    End If
    
    
    ' Receipt ID provided -> at this point it doesnt matter what shows up, just acknowledge it
    Dim popupCountCheck As Integer: popupCountCheck = 0
    
    Do
        popupCheckResult = PeopleSoft_Page_CheckForPopup(driver)
        If popupCheckResult.HasPopup = False Then Exit Do
        
        popupCountCheck = popupCountCheck + 1
        Debug.Print "Popup received after Receipt " & popupCountCheck & ": " & popupCheckResult.PopupText
        'rcpt.GlobalError = rcpt.GlobalError & "Popup Received after Receipt " & popupCountCheck & ": " & popupCheckResult.PopupText & vbCrLf
    
        If popupCheckResult.HasButtonYes Then ' Either Yes or OK....
            PeopleSoft_Page_AcknowledgePopup driver, popupCheckResult, vbYes
        Else
            PeopleSoft_Page_AcknowledgePopup driver, popupCheckResult, vbOK
        End If
        
        PeopleSoft_Page_WaitForProcessing driver
    Loop
        
    
    
    
    
    'PopupText = PeopleSoft_Page_SuppressPopup(driver, vbOK)
    'popUpIsExpected = InStr(1, PopupText, "This means the receipt is being updated by the receipt integration process") > 0  ' TODO: Change
    
    ' At this point, it doesnt matter if there is a popup
    'If False Then ' popUpIsExpected = False Then
    '    rcpt.HasGlobalError = True
    '    rcpt.GlobalError = "Unexpected popup: " & PopupText
        
    '    GoTo ReceiptFailed
    'End If
    
    
    PeopleSoft_PurchaseOrder_ProcessReceipt = True
    Exit Function
    
    
CancelReceiptAndExit:



ValidationFailed:
ReceiptFailed:
    PeopleSoft_PurchaseOrder_ProcessReceipt = False
    Exit Function
       
ExceptionThrown:
    PeopleSoft_PurchaseOrder_ProcessReceipt = False
    
    rcpt.HasGlobalError = True
    rcpt.GlobalError = "Exception: " & Err.Description
    



End Function
Public Function PeopleSoft_PurchaseOrder_RetrySaveWithBudgetCheck(ByRef session As PeopleSoft_Session, ByRef poRetryBC As PeopleSoft_PurchaseOrder_RetrySaveWithBudgetCheckParams) As Boolean
    
    
    On Error GoTo ExceptionThrown
    
    
    Dim By As New By, Assert As New Assert, Verify As New Verify
    Dim driver As New SeleniumWrapper.WebDriver
    
    
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



Public Function PeopleSoft_Page_SetValidatedField(ByRef driver As SeleniumWrapper.WebDriver, ByVal fieldElementID As String, ByVal fieldValue As String, ByRef fieldValResult As PeopleSoft_Field_ValidationResult, Optional ignoreEmptyValues As Boolean = True) As Boolean

     
    
    fieldValResult.ValidationFailed = False
    fieldValResult.ValidationErrorText = ""

    
        
    
    ' Dont bother if value is empty string or option to ignore empty values is false
    If Len(fieldValue) > 0 Or ignoreEmptyValues = False Then
        Dim elID As String, elVal As String
        
        'elID = fieldElement.getAttribute("id")
        elID = Replace(fieldElementID, "'", "\'")
        
        elVal = driver.executeScript("return document.getElementById('" & elID & "').value;")
        
        Dim tryNo As Integer
    
        tryNo = 1
        
        
        'Do

        If fieldValue <> elVal Then
            
            tryNo = 1
            
        
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
            
            fieldValResult.ValidationErrorText = PeopleSoft_Page_SuppressPopup(driver, vbOK)
            
            fieldValResult.ValidationFailed = fieldValResult.ValidationErrorText <> ""
            
            
            
            'driver.Wait 500
    
            elVal = driver.executeScript("return document.getElementById('" & elID & "').value;")
            
            
            If tryNo >= 3 Then
                fieldValResult.ValidationFailed = True
                fieldValResult.ValidationErrorText = "Could not set value"
            End If
            
            
            
        
        End If
    
            tryNo = tryNo + 1
        'Loop Until elVal <> ""

    End If
   
   
    
    'If Len(fieldValue) > 0 Then
    '
     '   pageFieldResult = PeopleSoft_Page_TypeCalculatedField(driver, fieldElement, fieldValue)
     '
     '
     '
     '   If pageFieldResult.alertPresent Then
     '       'purchaseOrder.HasError = True
     '       fieldValResult.ValidationFailed = True
     '       fieldValResult.ValidationErrorText = pageFieldResult.alertMsg
     '
     '   End If
    'End If
    
    PeopleSoft_Page_SetValidatedField = Not fieldValResult.ValidationFailed

End Function
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

    ' ---------------------------------------------------------------------
    ' Begin - Save w/ Budget Check
    ' ---------------------------------------------------------------------
    
    Dim By As New SeleniumWrapper.By
    
    
    Dim swByPOId As SeleniumWrapper.By
    Dim wePOId As SeleniumWrapper.WebElement
    
    Dim i As Integer
    
    driver.findElementById("PO_KK_WRK_PB_BUDGET_CHECK").Click
    
    PeopleSoft_Page_WaitForProcessing driver, TIMEOUT_LONG
    
    
    Dim PopupText As String
    
    PopupText = PeopleSoft_Page_SuppressPopup(driver, vbOK)
    
    If Len(PopupText) > 0 Then ' Error while saving
        budgetCheckResult.GlobalError = PopupText
        budgetCheckResult.HasGlobalError = True
        
        PeopleSoft_PurchaseOrder_SaveWithBudgetCheck = False
        Exit Function
    End If
    
    
    Set swByPOId = By.ID("Z_KK_ERR_WRK_PO_ID")
    
    Dim elementExists_budgetErrorId As Boolean
    
    elementExists_budgetErrorId = PeopleSoft_Page_ElementExists(driver, swByPOId)
    

    If Not elementExists_budgetErrorId Then
        'Set wePOId = driver.findElementById("PO_HDR_PO_ID$33$")
        Set wePOId = driver.findElementById("PO_HDR_PO_ID$14$") ' Fix for 2.9.1.1
         
        
        If wePOId.Text = "NEXT" Then ' Error while saving
            budgetCheckResult.GlobalError = "Unknown error - no PO ID"
            budgetCheckResult.HasGlobalError = True
            
            PeopleSoft_PurchaseOrder_SaveWithBudgetCheck = False
        
        Else
            budgetCheckResult.PO_ID = wePOId.Text
            PeopleSoft_PurchaseOrder_SaveWithBudgetCheck = True
        End If
        
        Exit Function
    End If
    
    Set wePOId = driver.findElement(swByPOId)
    'driver.findElement swBy.ID("ACE_Z_KK_ERR_WRK_")
    
    
    budgetCheckResult.PO_ID = wePOId.Text
    
    
    PeopleSoft_PurchaseOrder_BudgetCheckResult_ExtractFromPage driver, budgetCheckResult
    
    
    ' Click "Return"
    'driver.findElementById("PO_KK_WRK_PB_BUDGET_CHECK").Click
    
    
    
    PeopleSoft_PurchaseOrder_SaveWithBudgetCheck = True
    Exit Function
    ' ---------------------------------------------------------------------
    ' End - Save w/ Budget Check
    ' ---------------------------------------------------------------------
    
    
ErrOccured:
    PeopleSoft_PurchaseOrder_SaveWithBudgetCheck = False

End Function
Public Function PeopleSoft_PurchaseOrder_BudgetCheckResult_ExtractFromPage(driver As SeleniumWrapper.WebDriver, ByRef budgetCheckResult As PeopleSoft_PurchaseOrder_BudgetCheckResult) As Boolean

    Dim By As New SeleniumWrapper.By
    
    ' Click View All - by Line
    If PeopleSoft_Page_ElementExists(driver, By.ID("Z_KK_PO_ERR_VW$hviewall$0")) Then
        'driver.findElementById("Z_KK_PO_ERR_VW$hviewall$0").Click
        driver.runScript "javascript:submitAction_win0(document.win0,'Z_KK_PO_ERR_VW$hviewall$0');"
        PeopleSoft_Page_WaitForProcessing driver
    End If
    
    ' Click View All - by Project
    If PeopleSoft_Page_ElementExists(driver, By.ID("Z_KK_PRJ_ERR_VW$hviewall$0")) Then
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

    On Error GoTo findElementException:

    Dim we As SeleniumWrapper.WebElement
    
    Set we = driver.findElement(weBy, timeoutms)
    
    If Not we Is Nothing Then
        PeopleSoft_Page_ElementExists = True
        Exit Function
    End If
    
findElementException:

    PeopleSoft_Page_ElementExists = False
    

End Function
Private Function PeopleSoft_Page_GetElementText(driver As SeleniumWrapper.WebDriver, ByVal elementID As String) As String

    elementID = Replace(elementID, "'", "\'")

    PeopleSoft_Page_GetElementText = driver.executeScript("return document.getElementById('" & elementID & "').textContent;")

End Function

Public Sub PeopleSoft_Page_WaitForProcessing(driver As SeleniumWrapper.WebDriver, Optional timeout_s As Long = 60)

    
    Dim iter As Integer, procVisibility As Variant
    
    Const POLL_INTERVAL_MS As Double = 500 ' 0.5 s
    
    Dim loader_inProcess As Boolean
    
    
    'loader_inProcess = driver.executeScript("return (loader != null && loader.GetInProcess());")
    'Debug.Print "Loader_InProcess: " & loader_inProcess
    
    'Const TIMEOUT_MS As Double = (timeout_s * 1000 * (1000 / POLL_INTERVAL_MS))
    
    Dim MAX_ITER As Double
    
    MAX_ITER = timeout_s * 1000 / POLL_INTERVAL_MS
    
    iter = 0
    
    ' Processing is over when two actions happen (for good measure, both must occur):
    '   (1) The processing icon is no longer visible
    '   (2) When the PeopleSoft internal loader is no longer active and processing
    '
    Do
    
        loader_inProcess = driver.executeScript("return (loader != null && loader.GetInProcess());")
    
        procVisibility = driver.executeScript("return document.getElementById('WAIT_win0').style.visibility;")
    
        'Debug.Print "Visibility: " & ret
        
        driver.Wait POLL_INTERVAL_MS
        
        DoEvents
    
        iter = iter + 1

    Loop Until iter > MAX_ITER Or (procVisibility <> "visible" And loader_inProcess = False)
    
    
    
    
    
    If iter > MAX_ITER Then
        Err.Raise 513, , "Timeout"
    End If
    

End Sub
Public Function PeopleSoft_Page_CheckForPopup(driver As SeleniumWrapper.WebDriver) As PeopleSoft_Page_PopupCheckResult

    Dim popupCheckResult As PeopleSoft_Page_PopupCheckResult
    
    popupCheckResult.HasPopup = False



    On Error GoTo PopupNotFoundOrErr

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
    
    Set we = driver.findElementByXPath(".//*[@id='" & popupCheckResult.PopupElementID & "']/descendant::*[@id='alertmsg']/span")
    popupCheckResult.PopupText = we.Text
    
    
    popupCheckResult.HasButtonOk = PeopleSoft_Page_ElementExists(driver, By.XPath(".//*[@id='" & popupCheckResult.PopupElementID & "']/descendant::*[@id='#ICOK']"), 10)
    popupCheckResult.HasButtonCancel = PeopleSoft_Page_ElementExists(driver, By.XPath(".//*[@id='" & popupCheckResult.PopupElementID & "']/descendant::*[@id='#ICCancel']"), 10)
    popupCheckResult.HasButtonYes = PeopleSoft_Page_ElementExists(driver, By.XPath(".//*[@id='" & popupCheckResult.PopupElementID & "']/descendant::*[@id='#ICYes']"), 10)
    popupCheckResult.HasButtonNo = PeopleSoft_Page_ElementExists(driver, By.XPath(".//*[@id='" & popupCheckResult.PopupElementID & "']/descendant::*[@id='#ICNo']"), 10)
    
        
    PeopleSoft_Page_CheckForPopup = popupCheckResult
    
    Debug.Print "PeopleSoft_Page_CheckForPopup: ID='" & popupCheckResult.PopupElementID & "', " _
                & "Buttons=(" & IIf(popupCheckResult.HasButtonYes, "Yes", "") & IIf(popupCheckResult.HasButtonNo, "|No", "") & IIf(popupCheckResult.HasButtonOk, "|OK", "") & IIf(popupCheckResult.HasButtonCancel, "|Cancel", "") & "), " _
                & "Text='" & popupCheckResult.PopupText & "'"
    
    Exit Function
    
PopupNotFoundOrErr:
    popupCheckResult.HasPopup = False
    popupCheckResult.PopupElementID = ""
    popupCheckResult.PopupText = ""
    
    PeopleSoft_Page_CheckForPopup = popupCheckResult
    
    Debug.Print "PeopleSoft_Page_CheckForPopup: No popup found or error: " & Err.Description

End Function
Public Sub PeopleSoft_Page_AcknowledgePopup(driver As SeleniumWrapper.WebDriver, ByRef popupCheckResult As PeopleSoft_Page_PopupCheckResult, clickButton As VbMsgBoxResult)
    
    On Error GoTo ExceptionThrown
    
    If clickButton = vbOK Then
      driver.findElementByXPath(".//*[@id='" & popupCheckResult.PopupElementID & "']/descendant::*[@id='#ICOK']").Click
    ElseIf clickButton = vbCancel Then
      driver.findElementByXPath(".//*[@id='" & popupCheckResult.PopupElementID & "']/descendant::*[@id='#ICCancel']").Click
    ElseIf clickButton = vbYes Then
      driver.findElementByXPath(".//*[@id='" & popupCheckResult.PopupElementID & "']/descendant::*[@id='#ICYes']").Click
    ElseIf clickButton = vbNo Then
      driver.findElementByXPath(".//*[@id='" & popupCheckResult.PopupElementID & "']/descendant::*[@id='#ICNo']").Click
    End If
    
    
    Exit Sub
    
ExceptionThrown:
    Err.Raise Err.Number, Err.Source, "PeopleSoft_Page_AcknowledgePopup: " & Err.Description, Err.Helpfile, Err.HelpContext

End Sub
Public Function PeopleSoft_Page_SuppressPopup(driver As SeleniumWrapper.WebDriver, clickButton As VbMsgBoxResult, Optional matchText As String = "") As String

    Dim popupCheckResult As PeopleSoft_Page_PopupCheckResult

    On Error GoTo ExceptionThrown


    popupCheckResult = PeopleSoft_Page_CheckForPopup(driver)
    
    If popupCheckResult.HasPopup = False Then
        Debug.Print "PeopleSoft_Page_SuppressPopup: no popup found"
        Exit Function
    End If
    
    
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

Public Function PeopleSoft_Page_SuppressPopup_Old(driver As SeleniumWrapper.WebDriver, clickButton As VbMsgBoxResult) As String


On Error GoTo PopupNotFoundOrErr

    Dim we As SeleniumWrapper.WebElement
    Dim By As New SeleniumWrapper.By
    

    
    Dim wePopupModals As WebElementCollection
    Dim wePopupModal As WebElement
    Dim PopupText As String
    
    Dim popupModalContentID As String
    
    Set wePopupModals = driver.findElementsByXPath(".//*[contains(@id,'ptModContent_')]", 100)

    
    'Debug.Print wePopupModals.Count & " modals founds"
    
    If wePopupModals.Count = 0 Then Exit Function 'no popup modals found
    
    popupModalContentID = wePopupModals(0).getAttribute("id")
    
    'Debug.Print "modal content id: " & popupModalContentID

    
   
    'Set we = driver.findElementByXPath(".//*[@id='pt_modals']/descendant::*[@id='alertmsg']/span")
    'Set we = driver.findElementByXPath(".//*[@id='alertmsg']/span")
    Set we = driver.findElementByXPath(".//*[@id='" & popupModalContentID & "']/descendant::*[@id='alertmsg']/span")
    
    'Debug.Print "modal content text: " & we.Text
    PopupText = we.Text
    PeopleSoft_Page_SuppressPopup_Old = PopupText
    
    
    If clickButton = vbOK Then
        ' descendant of #okbutton
      'driver.findElementById("#ICOK").Click
      'driver.runScript "javascript:closeMsg('#ICOK');"
      driver.findElementByXPath(".//*[@id='" & popupModalContentID & "']/descendant::*[@id='#ICOK']").Click
    ElseIf clickButton = vbCancel Then
      'driver.findElementById("#ICCancel").Click
      driver.findElementByXPath(".//*[@id='" & popupModalContentID & "']/descendant::*[@id='#ICCancel']").Click
      'driver.runScript "javascript:aAction_win0(document.win0, '#ICCancel');"
      'driver.runScript "javascript:closeMsg('#ICCancel');"
    ElseIf clickButton = vbYes Then
      'driver.findElementById("#ICYes").Click
      driver.findElementByXPath(".//*[@id='" & popupModalContentID & "']/descendant::*[@id='#ICYes']").Click
      'driver.runScript "javascript:aAction_win0(document.win0, '#ICYes');"
      'driver.runScript "javascript:closeMsg('#ICYes);"
    End If
    
    
    
    
PopupNotFoundOrErr:
    Debug.Print "PeopleSoft_Page_SuppressPopup: " & PopupText

End Function
Private Function CurrencyFromString(strCur As String) As Currency

    strCur = Replace(strCur, ",", "")
    
    If IsNumeric(strCur) Then
        CurrencyFromString = CCur(strCur)
    Else
        CurrencyFromString = 0
    End If

End Function



