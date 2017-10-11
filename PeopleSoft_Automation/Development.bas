Attribute VB_Name = "Development"
Option Explicit
Option Private Module

Private Const Q_MAX_CONSECUTIVE_FAILURES = 6


Sub UpdateAppDisplay()
    
        Application.ScreenUpdating = True

End Sub

Public Sub Process_PO_ChangeOrder_Q_WorkInProgress()


    Dim i As Integer, j As Integer

    Dim queueTableRange As Range
    
    
    Dim ssMap As SPREADSHEET_MAP_2L
    
    Set queueTableRange = ThisWorkbook.Worksheets("PO_ChangeOrder_Q").Range("4:65536")
    Const HEADER_ROWS_SIZE = 0
    
    ' Column numbers (in order)
    Dim C_PO_BU As Integer
    Dim C_PO_ID As Integer
    Dim C_PO_DUE_DATE As Integer
    Dim C_PO_FLG_SEND_TO_VENDOR As Integer
    Dim C_CO_REASON As Integer
    Dim C_CO_STATUS As Integer
    Dim C_CO_PO_ERROR As Integer
    Dim C_CO_ITEM_ERROR As Integer
    
    Dim C_PO_LINE As Integer
    Dim C_PO_SCHEDULE As Integer

    
    Dim col As Integer
    col = 0
    col = col + 1: C_PO_BU = col
    col = col + 1: C_PO_ID = col
    col = col + 1: C_PO_LINE = col
    col = col + 1: C_PO_SCHEDULE = col
    col = col + 1: C_PO_DUE_DATE = col
    col = col + 1: C_PO_FLG_SEND_TO_VENDOR = col
    col = col + 1: C_CO_REASON = col
    col = col + 1: C_CO_STATUS = col
    col = col + 1: C_CO_PO_ERROR = col
    col = col + 1: C_CO_ITEM_ERROR = col

    
    
    
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
                PO_COs(idxParent).PO_ChangeOrder_Items(idxChild).PO_Line = -9999
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
        
            result = PeopleSoft_ChangeOrder_Process(session, PO_COs(idxParent))
            
            'Application.ScreenUpdating = False
            
            If result Then
                conseqfailCount = 0
            
                For idxChild = 1 To PO_COs(idxParent).PO_ChangeOrder_ItemCount
                    curRow = ssMap.PARENT_MAP(idxParent).MAP_CHILD_TO_ROW(idxChild)
                    
                    
                    queueTableRange.Cells(curRow, C_CO_STATUS).Value = "COMPLETE"
                Next idxChild
            Else
                ' -----------------------------------
                ' Begin - Build error strings and write to spreadsheet
                ' -----------------------------------
                Dim poErrString As String, itemErrString As String
                
                poErrString = ""
                
                If PO_COs(idxParent).PO_DUE_DATE_Result.ValidationFailed Then _
                    poErrString = poErrString & "|" & "PO_DUE_DATE: " & PO_COs(idxParent).PO_DUE_DATE_Result.ValidationErrorText & vbCrLf
                    
            
                For idxChild = 1 To PO_COs(idxParent).PO_ChangeOrder_ItemCount
                    itemErrString = ""
                    
                    If PO_COs(idxParent).PO_ChangeOrder_Items(idxChild).PO_Line > 0 Then
                        With PO_COs(idxParent).PO_ChangeOrder_Items(idxChild)
                            If .SCH_DUE_DATE_Result.ValidationFailed Then _
                                 itemErrString = itemErrString & "|" & "SCH_DUE_DATE: " & .SCH_DUE_DATE_Result.ValidationErrorText & vbCrLf
                        End With
                    End If
            
                    curRow = ssMap.PARENT_MAP(idxParent).MAP_CHILD_TO_ROW(idxChild)
                    
                    queueTableRange.Cells(curRow, C_CO_STATUS).Value = "<ERROR>"
                    queueTableRange.Cells(curRow, C_CO_PO_ERROR).Value = PO_COs(idxParent).GlobalError & IIf(Len(poErrString) > 0, vbCrLf & poErrString, "")
                    queueTableRange.Cells(curRow, C_CO_ITEM_ERROR).Value = itemErrString
                    
                    
                    queueTableRange.Cells(curRow, C_CO_PO_ERROR).WrapText = False
                    queueTableRange.Cells(curRow, C_CO_ITEM_ERROR).WrapText = False
                Next idxChild
                ' -----------------------------------
                ' End - Build error strings and write to spreadsheet
                ' -----------------------------------
                
                conseqfailCount = conseqfailCount + 1
                
                If conseqfailCount > Q_MAX_CONSECUTIVE_FAILURES Then Exit Sub
            End If
            
            
            'Application.ScreenUpdating = True
            
            Debug.Print
        
        End If
    Next idxParent


    
End Sub

