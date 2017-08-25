Attribute VB_Name = "JH_CommonVBA"
Option Explicit
Option Private Module

' JH_Common_VBA
' ------------------------------------------------------------------------------------------------------------------
' Common VBA functions used across multiple VBA projects
'
' Joseph Huntley (joseph.huntley@vzw.com)
' ------------------------------------------------------------------------------------------------------------------
' Changelog:
'
' 2017-08-24 (joe h)    - Added Workbook_SaveCopyWithoutMacros
'                       - Added Workbook_RefreshPivotTables
'                       - Added Clipboard_CopyText
'                       - Converted to a private module so public functions do not show up in macro lists
' 2017-08-22 (joe h)    - Added DataTable_InsertBlankRows, DataTable_AppendWorksheet
'                       - Added Worksheet_ColumnFixNumbersStoredAsText
'                       - Added Worksheet_SetColumnFormats
'                       - Worksheet_InsertColumnUsingTemplate: Fixed bug:
' 2017-03-31 (joe h)    - Renamed SpreadsheetTableToPivotMap -> Table_GroupBy_Map and added comments
' 2016-10-28 (joe h)    - Added Token_Extract_WIP
' 2016-06-31 (joe h)    - Added RegEx_Replace
'                       - Renamed GetTableRange to Range_FromAny
'                       - Renamed Worksheet_LoadColumnDataToArray to Data_LoadFromRange
' 2016-05-31 (joe h)    - Added Array_ResizeOverflow, RegEx_Extract
' 2016       (joe h)    - Migrated URL_BuildQueryString, URL_Encode, URL_Decode
'                       - Added Output_XXX, Debug_XXX functions
' 2015-12-28 (joe h)    - Added SpreadsheetTableToPivotMap (and dependencies)
'                       - Added Worksheet_GetColumnsByName, Worksheet_GetColumnDictionary
' 2015-06-30 (joe h)    - Added Worksheet_HasAllColumns
' 2015-06-12 (joe h)    - Added Worksheet_Append
' 2015-05-13 (joe h)    - Added GetTableRange, Lookup_MultiMatch
' 2015-05-05 (joe h)    - Worksheet_InsertColumnUsingTemplate: Added ability to use formulas
'                       - Worksheet_ColData: Fixed to return array instead of range
' 2015-05-01 (joe h)    - Array_AppendMultiple: Fixed error when array is non-initialized
'                       - Array_CreateAndFill: Fixed error when numItems = 0
' 2015-04-17 (joe h)    - Replaced early bound WinHttpRequest object with late bound version (multiple procedures)
' 2015-03-25 (joe h)    - Updated Worksheet_GetLastRow/Worksheet_GetLastColumn with new formula
'                       - Added function Range_LastColumn
' ------------------------------------------------------------------------------------------------------------------



#If VBA7 Then
    Private Declare PtrSafe Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
    Private Declare PtrSafe Function GetTempFileName Lib "kernel32" Alias "GetTempFileNameA" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal lngMilliSeconds As Long)
#Else
    Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
    Private Declare Function GetTempFileName Lib "kernel32" Alias "GetTempFileNameA" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long
    Private Declare Sub Sleep Lib "kernel32" (ByVal lngMilliSeconds As Long)
#End If


' Deprecated since 2017-03-31. Use Table_GroupBy_Map() variant instead
Private Type SpreadsheetTableToPivotMap_UniquePivotValue_Old
    ID As String
    Value As Variant
    PivotCol As Long
    ParentIndex As Long
    RelativeIndex As Long
    ChildCount As Long
    RowNumbers As Collection
    FirstRowInContext As Long
End Type


' Deprecated since 2017-03-31. Use Table_GroupBy_Map() variant instead
Public Type SpreadsheetTableToPivotMapping_Old
    PivotToRowMap() As Variant
    RowToPivotMap() As Long
End Type


Public Enum TableGroupByMap_ZeroIndexItemType
    Group_Undefined = 0
    Group_ItemCount = 1
    Group_Value = 2
End Enum

Public Type TableGroupByMapping
    GroupToRowMap() As Variant
    RowToGroupMap() As Long
End Type


Private Type TableGroupByMap_UniqueGroupValue
    ID As String
    Value As Variant
    GroupCol As Long
    ParentIndex As Long
    RelativeIndex As Long
    ChildCount As Long
    RowNumbers As Collection
    FirstRowInContext As Long
End Type



Private Excel_AppUpdates_prevCalculation As XlCalculation
Private Excel_AppUpdates_prevScreenUpdating As Boolean
Private Excel_AppUpdates_callCount As Integer


Private Debug_CurLine As Long
Private Debug_Lines() As String
Private Debug_Capacity As Long
Private Debug_SilentFlag As Boolean

Private Output_CurLine As Long
Private Output_Lines() As String
Private Output_Capacity As Long
Private Output_SilentFlag As Boolean
Private Output_IndentLevel As Integer
Private Output_CM_CurLines() As Long
Private Output_CM_Lines() As String
Private Output_CM_CapacityPerCol As Long
Private Output_CM_NumCols As Integer
Private Output_CM_HBar_Lines() As Integer
Private Output_CM_HBar_Count As Integer

Public Sub Excel_AppUpdates_Reset()

    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    
    Excel_AppUpdates_callCount = 0

End Sub
Public Sub Excel_AppUpdates_Disable()
    
    If Excel_AppUpdates_callCount = 0 Then
        Excel_AppUpdates_prevScreenUpdating = Application.ScreenUpdating
        Excel_AppUpdates_prevCalculation = Application.Calculation
    
        
        Application.ScreenUpdating = False
        Application.Calculation = xlCalculationManual
    End If

    Excel_AppUpdates_callCount = Excel_AppUpdates_callCount + 1

End Sub
Public Sub Excel_AppUpdates_Restore()

    Excel_AppUpdates_callCount = Excel_AppUpdates_callCount - 1
    
    If Excel_AppUpdates_callCount = 0 Then
        Application.ScreenUpdating = Excel_AppUpdates_prevScreenUpdating
        Application.Calculation = Excel_AppUpdates_prevCalculation
    End If

End Sub
' Copy text to clipboard
Public Sub Clipboard_CopyText(text As String)

    Dim dataObj As Object
    
    Set dataObj = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
    
    dataObj.SetText text
    dataObj.PutInClipboard
    
    Set dataObj = Nothing
    
End Sub

' CreateTempFileName: returns an unused temporary filename in the %TEMP% directory with the specified prefix
'
' 2017-08-22: Converted to public function
Public Function CreateTempFileName(prefix As String) As String
   Dim tempPath As String * 512
   Dim tempName As String * 576
   Dim nRet As Long

   nRet = GetTempPath(512, tempPath)
   If (nRet > 0 And nRet < 512) Then
      nRet = GetTempFileName(tempPath, prefix, 0, tempName)
      
      If nRet <> 0 Then CreateTempFileName = Left$(tempName, InStr(tempName, vbNullChar) - 1)
   End If
   
End Function
Public Sub WaitSeconds(intSeconds As Integer)
  ' Comments: Waits for a specified number of seconds
  ' Params  : intSeconds      Number of seconds to wait
  ' Source  : Total Visual SourceBook

  On Error GoTo PROC_ERR

  Dim datTime As Date

  datTime = DateAdd("s", intSeconds, Now)

  Do
   ' Yield to other programs (better than using DoEvents which eats up all the CPU cycles)
    Sleep 100
    DoEvents
  Loop Until Now >= datTime

PROC_EXIT:
  Exit Sub

PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , "WaitSeconds"
  Resume PROC_EXIT
End Sub

Private Function URL_Encode(stringVal As String, Optional spaceAsPlus As Boolean = False) As String

  Dim StringLen As Long: StringLen = Len(stringVal)

  If StringLen > 0 Then
    ReDim result(StringLen) As String
    Dim i As Long, CharCode As Integer
    Dim char As String, Space As String

    If spaceAsPlus Then Space = "+" Else Space = "%20"

    For i = 1 To StringLen
      char = Mid$(stringVal, i, 1)
      CharCode = Asc(char)
      Select Case CharCode
        Case 97 To 122, 65 To 90, 48 To 57, 45, 46, 95, 126
          result(i) = char
        Case 32
          result(i) = Space
        Case 0 To 15
          result(i) = "%0" & Hex(CharCode)
        Case Else
          result(i) = "%" & Hex(CharCode)
      End Select
    Next i
    
    URL_Encode = Join(result, "")
  End If
  
  
End Function
Private Function URL_Decode(stringToDecode As String) As String
    
    Dim tempAns As String
    Dim curChr As Integer
    
    curChr = 1
    
    Do Until curChr - 1 = Len(stringToDecode)
        Select Case Mid(stringToDecode, curChr, 1)
            Case "+"
                tempAns = tempAns & " "
            Case "%"
                tempAns = tempAns & Chr(val("&h" & Mid(stringToDecode, curChr + 1, 2)))
                curChr = curChr + 2
            Case Else
                tempAns = tempAns & Mid(stringToDecode, curChr, 1)
        End Select
        
        curChr = curChr + 1
    Loop
    
    URL_Decode = tempAns

End Function
Public Function URL_BuildQueryString(ParamArray fieldPairs() As Variant) As String

    Dim queryString As String
    Dim i As Integer
    
    
    Dim fieldPairCount As Integer
    
    fieldPairCount = IIf(UBound(fieldPairs) > -1, UBound(fieldPairs) - LBound(fieldPairs) + 1, 0)
    If fieldPairCount Mod 2 <> 0 Then Err.Raise -1, , "URL_BuildQueryString: Arguments must be multiples of two (field name and value pairs)."
    
    
    Dim fieldData As Variant, fieldValue As Variant
    
    Dim fieldName As String
    
    i = 1
    
    For Each fieldData In fieldPairs
        
        If i Mod 2 = 1 Then  ' Odd -> Field Data represents a Field Name
            fieldName = URL_Encode(CStr(fieldData))
        Else    ' Even -> Field Data represents a Field Value
            If IsArray(fieldData) Then  ' Field Name with multiple values
                For Each fieldValue In fieldData
                    queryString = queryString & "&" & fieldName & "=" & URL_Encode(CStr(fieldValue))
                Next fieldValue
            Else
            
                queryString = queryString & "&" & fieldName & "=" & URL_Encode(CStr(fieldData))
            End If
        End If
        
    
        i = i + 1
    Next fieldData
    
    ' Remove first & at beginning of query string
    If Len(queryString) > 0 Then queryString = Right$(queryString, Len(queryString) - 1)
    
    URL_BuildQueryString = queryString
    

End Function

Public Function RegEx_Replace(str As String, pattern As String, replaceStr As String) As Variant
    
    ' joe h:    Added 6/13/16
    
    Dim regExp As Object 'New VBScript_RegExp_55.regExp
    Dim regExpMatches As Object 'VBScript_RegExp_55.MatchCollection
    
    Set regExp = CreateObject("VBScript.RegExp")
    
    regExp.Global = True
    regExp.pattern = pattern
    
    RegEx_Replace = regExp.Replace(str, replaceStr)
    
    
    Set regExp = Nothing
    Set regExpMatches = Nothing
    

End Function
' RegEx_Extract: Extract part of a string using a specific pattern. Shortcut utility function for calling regex objects.
'   Returns array with values of each regexp group within the string
'
Public Function RegEx_Extract(str As String, pattern As String, Optional delim As String = "|") As Variant

    ' joe h:    Added 5/31/16
    


    Dim retArr() As Variant
    Dim i As Long, grp_count As Long

    Dim regExp As Object 'New VBScript_RegExp_55.regExp
    Dim regExpMatches As Object 'VBScript_RegExp_55.MatchCollection
    
    Set regExp = CreateObject("VBScript.RegExp")
    
    
    regExp.Global = True
    regExp.pattern = pattern
    
    Set regExpMatches = regExp.Execute(str)
    
    grp_count = regExpMatches(0).SubMatches.count
    ReDim retArr(1 To grp_count)

    For i = 1 To grp_count
        retArr(i) = regExpMatches(0).SubMatches(i - 1)
    Next i
    
    Set regExpMatches = Nothing
    Set regExp = Nothing
    
    'If grp_count = 1 Then
    '    RegEx_Extract = retArr(1)
    '    Exit Function
    'End If
    
    
    ' If calling directly from excel as an array formula, then modify the result to fit the cells
    If IsObject(Application.Caller) Then
        Dim callerRows As Long, callerCols As Long
    
        callerRows = Application.Caller.rows.count
        callerCols = Application.Caller.columns.count
        
        If callerRows = 1 And callerCols = 1 Then
            RegEx_Extract = retArr(1)
            Exit Function
        End If
        
        retArr = Array_ResizeOverflow(retArr, callerRows, callerCols)
    End If
    
    RegEx_Extract = retArr
    
End Function


Public Function GetWebData(url As String, destSheetName As String, Optional appendToSheet As Boolean = False) As Boolean

    ' MODIFIED

    ' GetWebData Macro, populate destSheetName with  data
    
    GetWebData = False
    
    'Clear clipboard for better performance
    Application.CutCopyMode = False
    Application.CutCopyMode = True
    
    Application.DisplayAlerts = False
    
    
    Dim ws As Worksheet, qryTableName As String
    Dim destRng As Range
     
    
    'Clear current sheet, or add a sheet if it doesnt exist
    If Not Worksheet_SheetExists(destSheetName) Then
        Set ws = ThisWorkbook.Sheets.Add
        ws.Name = destSheetName
    Else
        Set ws = ThisWorkbook.Sheets(destSheetName)
        
        If Not appendToSheet Then ws.Cells.ClearContents
    End If
    
    
    If appendToSheet Then
        Dim i As Integer
        
        i = 1
        Do While ws.Cells(i, 1) <> "": i = i + 1: Loop
        
        Set destRng = ws.Cells(i, 1)
    Else
        Set destRng = ws.Cells(1, 1)
    End If
    
        
    
    Application.DisplayAlerts = True
    
    
    qryTableName = "GetWebData_" & Format(Now(), "yyyymmddhhmmss")
    
    
    Dim conn As String
    
    If False Then
    If Len(url) > 127 Then
        Dim file As String
        
        file = GetWebData_DownloadFile(url, , "")
        
        If file = "" Then
            Err.Raise -1, , "Downloading file failed: URL: " & url
        End If
        
        
        conn = "URL;" & file
    Else
        conn = "URL;" & url
    End If
    End If
    
    
        conn = "URL;" & url
    
    'Grab  data, paste into destSheet
    With ws.QueryTables.Add(Connection:= _
        conn, Destination:=destRng)
        .Name = qryTableName
        .FieldNames = True
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .BackgroundQuery = True
        .RefreshStyle = XlCellInsertionMode.xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0

        .WebSelectionType = XlWebSelectionType.xlEntirePage
        .WebFormatting = XlWebFormatting.xlWebFormattingNone
        .WebPreFormattedTextToColumns = True
        .WebConsecutiveDelimitersAsOne = True
        .WebSingleBlockTextImport = False
        .WebDisableDateRecognition = False
        .WebDisableRedirections = False
        
        .Refresh BackgroundQuery:=False
        
        .Delete
    End With
    
    'ws.Names(qryTableName).Delete 'clean up
    
    GetWebData = True
            
End Function

Function GetWebData_CSV(url As String, destSheetName As String, Optional appendToSheet As Boolean = True) As Boolean

    ' GetWebData_CSV Macro, populate destSheetName with CSV data
    
    GetWebData_CSV = False
    
    'Clear clipboard for better performance
    Application.CutCopyMode = False
    Application.CutCopyMode = True
    
    Application.DisplayAlerts = False
    
    
    Dim ws As Worksheet, qryTableName As String
    Dim qt As QueryTable
    Dim destRng As Range
     
    
    'Clear current sheet, or add a sheet if it doesnt exist
    If Not Worksheet_SheetExists(destSheetName) Then
        Set ws = ThisWorkbook.Sheets.Add
        ws.Name = destSheetName
    Else
        Set ws = ThisWorkbook.Sheets(destSheetName)
        
        If Not appendToSheet Then ws.Cells.ClearContents
    End If
    
    
    If appendToSheet Then
        Dim lastRow As Long
        
        lastRow = Worksheet_GetLastRow(ws)
        
        Set destRng = ws.Cells(lastRow + 1, 1)
    Else
        Set destRng = ws.Cells(1, 1)
    End If
    
        
    
    Application.DisplayAlerts = True
    
    
    qryTableName = "GetWebData_" & Format(Now(), "yyyymmddhhmmss")

    Dim conn As String


    ' some funky things happen with long URLs. if the URL is long, then download first and then open
    If Len(url) > 127 Then
        Dim csvFile As String
        
        csvFile = GetWebData_DownloadFile(url, , "csv")
        
        If csvFile = "" Then
            Err.Raise -1, , "Downloading file failed: URL: " & url
        End If
        
        
        conn = "TEXT;" & csvFile
    Else
        conn = "TEXT;" & url
    End If
    
    
    
    
    With ws.QueryTables.Add(Connection:=conn, Destination:=destRng)
        .Name = qryTableName
                
        .FieldNames = True
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .BackgroundQuery = True
        .RefreshStyle = XlCellInsertionMode.xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        
        .TextFilePromptOnRefresh = False
        .TextFilePlatform = XlPlatform.xlWindows
        .TextFileStartRow = 1
        .TextFileParseType = XlTextParsingType.xlDelimited
        .TextFileTextQualifier = XlTextQualifier.xlTextQualifierDoubleQuote
        .TextFileConsecutiveDelimiter = False
        .TextFileTabDelimiter = True
        .TextFileSemicolonDelimiter = False
        .TextFileCommaDelimiter = True
        .TextFileSpaceDelimiter = False
        .TextFileColumnDataTypes = Array(XlColumnDataType.xlGeneralFormat)
        .TextFileDecimalSeparator = ","
        .TextFileTrailingMinusNumbers = True
        
        .Refresh BackgroundQuery:=False
        
        .Delete
    End With
    
    'ws.Names(qryTableName).Delete 'clean up
    
    GetWebData_CSV = True
            
End Function
Function GetWebData_DownloadFile(url As String, Optional fileName As String, Optional fileExt As String, Optional winHttpReq As Object = Nothing) As String

    'Dim url As String: url = "http://vaculptpa13.nss.vzwnet.com:8282/alte/createcsv?fileName=http://vaculptpa31.nss.vzwnet.com:8282/reports/alte/w503686/sah-w503686-Thu_Nov_20_17_38_14_EST_2014_11202014173818.xls"
    'Dim fileName As String: fileName = "C:\test7.csv"
    'Dim fileExt As String: fileExt = "csv"
    
    'Dim winHttpReq As Object: Set winHttpReq = Nothing



    If Len(fileName) = 0 Then fileName = CreateTempFileName("VBA") & IIf(Len(fileExt) > 0, "." & fileExt, "")


    Dim httpReq As Object 'WinHttpRequest
    
    If winHttpReq Is Nothing Then
        Set httpReq = CreateObject("WinHttp.WinHttpRequest.5.1")
        httpReq.SetTimeouts 0, -1, -1, -1 ' do not timeout
    Else
        Set httpReq = winHttpReq
    End If
    

   
    httpReq.Open "GET", url, False
    httpReq.Send

    
    If httpReq.Status <> 200 Then
        Set httpReq = Nothing
        GetWebData_DownloadFile = ""
        Exit Function
    End If
    
    
    Dim responseData As Variant
    responseData = httpReq.ResponseBody
    
    Set httpReq = Nothing
    
    Dim oStream As Object
    Set oStream = CreateObject("ADODB.Stream")
    
    With oStream
        .Type = 1
        .Open
        .Write responseData
        .SaveToFile fileName, 2
        .Close
    End With
    
    Set oStream = Nothing
    
    
    Debug.Print "GetWebData_DownloadFile: " & fileName
    
    GetWebData_DownloadFile = fileName
    
    
    'Dim hFile As Integer
    
    'hFile = FreeFile
    
    'Open fileName For Binary Access Write As #hFile
    '    Put #hFile, 1, responseData
    'Close #hFile
    
    
    
    
    Exit Function
    
FileErrorOccurred:
    'Close #hFile
        

End Function

Function Worksheet_GetHeaderColNum(colRng As Range, colName As String) As Integer


    Dim findR As Range
    
    
    Set findR = colRng.Find(colName, LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlNext, matchCase:=True)
    
    If findR Is Nothing Then
        Worksheet_GetHeaderColNum = 0
    Else
        Worksheet_GetHeaderColNum = findR.column
    End If
    

End Function
Function Worksheet_GetColumnByName(ws As Worksheet, colName As String, Optional headerRow As Long = 1, Optional caseSensitive As Boolean = True) As Integer

    Dim r As Range, findR As Range
    
    Set r = ws.Cells
    
    If headerRow > 1 Then r = r.Offset(headerRow)
    
    Set findR = r.Find(colName, LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlNext, matchCase:=caseSensitive)
    
    If findR Is Nothing Then
        Worksheet_GetColumnByName = 0
    Else
        Worksheet_GetColumnByName = findR.column
    End If
    

End Function


' Recursively resolves columns names within an array to its column numbers. Existing numbers are left alone
Public Function Worksheet_GetColumnsByName(ws As Worksheet, columns As Variant, Optional headerRow As Long = 1) As Variant ' Recursive

    Dim ret As Variant
    Dim i As Long
    
    If IsArray(columns) Then
        ret = columns
        
        For i = LBound(columns) To UBound(columns)
            ret(i) = Worksheet_GetColumnsByName(ws, columns(i), headerRow)
        Next i
        
    Else
        If IsNumeric(columns) Then
            ret = columns
        Else
            ret = Worksheet_GetColumnByName(ws, CStr(columns), headerRow)
        End If
    End If
    
    Worksheet_GetColumnsByName = ret

End Function
Public Function Worksheet_GetColumnDictionary(ws As Worksheet, Optional headerRow As Long = 1) As Object


    'Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("Sheet1")
    'Dim headerRow As Long: headerRow = 1


    Dim dict As Object 'Dictionary
    Dim lastCol As Long, origColName As String, tryColName As String
    Dim i As Long, j As Long
    
    Set dict = CreateObject("Scripting.Dictionary") 'New Dictionary
    
    lastCol = Worksheet_GetLastColumn(ws)
    
    For i = 1 To lastCol
        origColName = ws.Cells(headerRow, i)
        tryColName = origColName
    
        j = 2
        Do While dict.Exists(tryColName) ' Try ColName, ColName2, ColName3 to avoid duplicates
            tryColName = origColName & "_" & j
            j = j + 1
        Loop
        
        dict.Add Key:=tryColName, Item:=i
    Next i
    
    
    Set Worksheet_GetColumnDictionary = dict
    
    Set dict = Nothing

End Function
Function Worksheet_HasAllColumns(sheet As Variant, ParamArray colNames() As Variant) As Boolean
    
    Dim ws As Worksheet
    
    Worksheet_HasAllColumns = False
    
    
    If TypeName(sheet) = "Worksheet" Then
        Set ws = sheet
    Else
        If Worksheet_SheetExists(CStr(sheet)) Then
            Set ws = ThisWorkbook.Sheets(sheet)
        Else
            Err.Raise -1, , "Invalid sheetname: " & sheet
            Exit Function
        End If
    End If
    
    Dim col As Variant
    
    For Each col In colNames
        If Worksheet_GetColumnByName(ws, CStr(col)) = 0 Then Exit Function
    Next col
    
    
    Worksheet_HasAllColumns = True

End Function
Function Worksheet_GetColumnAsRange(wsRng As Range, colName As String, Optional headerRow As Long = 1) As Range

    Dim r As Range, findR As Range
    
    
    If headerRow > 1 Then wsRng = wsRng.Offset(headerRow)
    
    Set findR = wsRng.Find(colName, LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlNext, matchCase:=False)
    
    If findR Is Nothing Then
        Worksheet_GetColumnAsRange = CVErr(xlErrNA)
    Else
        Set Worksheet_GetColumnAsRange = wsRng.columns(findR.column)
    End If
    

End Function
Public Function Worksheet_ColData(sheetName As String, colName As String, Optional includeHeaders As Boolean = True, Optional headerRow As Long = 1) As Variant

    
    

    'On Error GoTo ErrHandler:

    Dim ws As Worksheet
    Dim lastRow As Long, colNbr As Long
    
    If Not Worksheet_SheetExists(sheetName) Then
        Worksheet_ColData = CVErr(xlErrName)
        Exit Function
    End If
        
    Set ws = ThisWorkbook.Sheets(sheetName)
    
    colNbr = Worksheet_GetColumnByName(ws, colName, headerRow)
    
    If colNbr = 0 Then
        Worksheet_ColData = CVErr(xlErrName)
        Exit Function
    End If
    
    
    
    
    lastRow = Worksheet_GetLastRow(ws, headerRow)

    Dim rng As Range, data As Variant
    
    
    Set rng = ws.Cells(headerRow + IIf(includeHeaders, 0, 1), colNbr).Resize(lastRow + IIf(includeHeaders, 0, -1))
    
    data = rng
    Worksheet_ColData = data
    

    
    
    Exit Function
    
'ErrHandler:
 '   Worksheet_ColData = CVErr(xlErrValue)
    

End Function


Public Function Worksheet_GetLastColumn(ws As Worksheet) As Long
    On Error Resume Next
    Worksheet_GetLastColumn = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), LookAt:=xlPart, LookIn:=xlFormulas, _
                            SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, matchCase:=False).column
    On Error GoTo 0
End Function
Public Function Worksheet_GetLastRow(ws As Worksheet, Optional startRow As Long = 1) As Long
    On Error Resume Next
    Worksheet_GetLastRow = ws.Cells.Find(What:="*", After:=ws.Cells(startRow, 1), LookAt:=xlPart, LookIn:=xlFormulas, _
                        SearchOrder:=xlByRows, SearchDirection:=xlPrevious, matchCase:=False).row
    On Error GoTo 0
End Function

Function Worksheet_SheetExists(sheetName As String, Optional wb As Workbook) As Boolean
    
    ' 2015-06-12 (joe h): Added parameter wb

    ' returns TRUE if the sheet exists in the active workbook
    Worksheet_SheetExists = False
    
    On Error GoTo NoSuchSheet
    
    If IsMissing(wb) Or wb Is Nothing Then
        Set wb = ThisWorkbook
    End If
    
    If Len(wb.Sheets(sheetName).Name) > 0 Then
        Worksheet_SheetExists = True
        Exit Function
    End If
    
NoSuchSheet:

End Function

Public Sub Worksheet_FixNumbersStoredAsText(ws As Worksheet)


    Dim lastCol As Long, lastRow As Long
    Dim i As Long
    
    
    lastCol = Worksheet_GetLastColumn(ws)
    lastRow = Worksheet_GetLastRow(ws)


    Dim appCalc As XlCalculation
    Dim appScreenUpdate As Boolean

    appCalc = Application.Calculation
    appScreenUpdate = Application.ScreenUpdating
    
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    
    Dim rng As Range
    
    Dim data() As Variant
    
    
    Set rng = ws.Cells(1, 1).Resize(lastRow, lastCol)
    rng.NumberFormat = "General"
    data = rng.Value
    rng.Value = data
    
    'For i = 1 To lastCol
    '    Set rng = ws.Cells(1, i).Resize(lastRow)
    '    rng.NumberFormat = "General"
    '    rng.Value = rng.Value
    'Next i


    Application.Calculation = appCalc
    Application.ScreenUpdating = appScreenUpdate
    


End Sub
Public Sub Worksheet_ColumnFixNumbersStoredAsText(ws As Worksheet, column As Variant, Optional headerRow As Long = 1)


    Dim lastRow As Long, colNbr As Long
    Dim i As Long
    
    
    lastRow = Worksheet_GetLastRow(ws)
    
    If IsNumeric(column) Then
        colNbr = CLng(column)
    Else
        colNbr = Worksheet_GetColumnByName(ws, CStr(column))
    End If


    Dim appCalc As XlCalculation
    Dim appScreenUpdate As Boolean

    appCalc = Application.Calculation
    appScreenUpdate = Application.ScreenUpdating
    
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    
    Dim rng As Range
    
    Set rng = ws.Cells(headerRow, i).Resize(lastRow - headerRow + 1)
    rng.NumberFormat = "General"
    rng.Value = rng.Value


    Application.Calculation = appCalc
    Application.ScreenUpdating = appScreenUpdate
    


End Sub
' Worksheet_SetColumnFormats: Sets column number formats in a worksheet by column names
'    Wildcards are supported (multiple columns)
'
' Examples:
'     1) Set Order Date column format to date, and Qty column to a number
'       Worksheet_SetColumnFormat(Sheets("MySheet"), "Order Date", "m/d/yy", "Qty", "0")
'     2) Wildcard usage: Set all columns ending with 'Date' to be formatted as a date
'       Worksheet_SetColumnFormat(Sheets("MySheet"), "*Date", "m/d/yy")
Public Sub Worksheet_SetColumnFormats(ws As Worksheet, ParamArray columnsAndFormats() As Variant)

    Dim i As Integer
    
    
    Dim paramCount As Integer
    
    paramCount = IIf(UBound(columnsAndFormats) > -1, UBound(columnsAndFormats) - LBound(columnsAndFormats) + 1, 0)
    If paramCount Mod 2 <> 0 Then Err.Raise -1, , "Worksheet_SetColumnFormat: Arguments must be multiples of two (column name and format pairs)."
    
    Dim columnFormatData As Variant
    Dim columnName As String, columnFormat As String
    
    Dim r As Range, findR As Range, firstAddressFound As String
    
    
    
    'If headerRow > 1 Then r = r.Offset(headerRow)
    i = 1
    
    For Each columnFormatData In columnsAndFormats
        
        If i Mod 2 = 1 Then
            ' Odd -> Data represents a Column Name (first part of pair)
            columnName = CStr(columnFormatData)
        Else
            ' Even -> Data represents column type value (second part of pair)
            columnFormat = CStr(columnFormatData)
            
            Set r = ws.Cells.rows(1)
            Set findR = r.Find(columnName, LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlNext, matchCase:=False)
            
            If Not findR Is Nothing Then
                firstAddressFound = findR.Address
                
                Do
                    findR.EntireColumn.NumberFormat = columnFormat
                
                    Set findR = r.FindNext(findR)
                Loop While Not findR Is Nothing And findR.Address <> firstAddressFound
            End If
            
        End If
        
        i = i + 1
    Next columnFormatData
    
    
    

End Sub

Sub Worksheet_FillColumn_FromRange(destSheet As Worksheet, destColumn As Integer, sourceRange As Range, Optional destStartRow As Long, Optional sourceCol As Integer = 1)

    If IsMissing(destStartRow) Then destStartRow = Worksheet_GetLastRow(destSheet) + 1
    
    Dim i As Integer
    
    For i = 1 To sourceRange.rows.count
        destSheet.Cells(destStartRow + i - 1, destColumn).Value = sourceRange.Cells(i, sourceCol).Value
    Next i
    

End Sub
Sub Worksheet_FillColumn_ByValue(destSheet As Worksheet, destColumn As Integer, val As Variant, numRows As Long, Optional destStartRow As Long)

    If IsMissing(destStartRow) Then destStartRow = Worksheet_GetLastRow(destSheet) + 1
    
    Dim i As Integer
    
    For i = 1 To numRows
        destSheet.Cells(destStartRow + i - 1, destColumn).Value = val
    Next i

End Sub


' Worksheet_InsertColumnUsingTemplate
' ----------------------------------------------------------
' Inserts calculated column which refers to other columns values by name
'
' Example:
'   Sheet1 has two columns A,B with heading "foo", "bar"
'
'   Add new column "car" which contains the content of column foo and bar separated by a dash (Formula: =$Ax&"-"&$Bx):
'       Worksheet_InsertColumnUsingTemplate(Sheets("Sheet1"), "car", "{foo}-{bar}", False)
'
'   As a formula: Add new column "test" which contains the content of sum of foo and bar, formatted using TEXT() function (Formula: ="TEXT($Ax + $Bx, ""00.00"")"):
'       Worksheet_InsertColumnUsingTemplate(Sheets("Sheet1"), "test", "TEXT({foo} + {bar}, ""00.00"")", True)
'
Public Sub Worksheet_InsertColumnUsingTemplate(sheet As Worksheet, newColName As String, colTemplate As String, Optional insertAfterColumn As Variant, Optional isFormula As Boolean = False, Optional isArrayFormula As Boolean = False, Optional colDelimPrefix As String = "{", Optional colDelimSuffix = "}")




            
    Dim insertAfterColumnNum As Integer
    
    
    If IsMissing(insertAfterColumn) Then
        insertAfterColumnNum = Worksheet_GetLastColumn(sheet)
    Else
        If IsNumeric(insertAfterColumn) = False Then
            insertAfterColumnNum = Worksheet_GetColumnByName(sheet, CStr(insertAfterColumn))
            
            If insertAfterColumnNum = 0 Then Err.Raise -1, , "Cannot find column '" & insertAfterColumn & "' in worksheet " & sheet.Name
        Else
            insertAfterColumnNum = insertAfterColumn
        End If
    End If
    
    If insertAfterColumnNum <= 0 Then insertAfterColumnNum = Worksheet_GetLastColumn(sheet)
    
    
    
    
    Dim ret As String
    Dim curPos As Long, nextColNameStartPos As Long, nextColNameEndPos
    Dim token As String
    Dim colName As String, colNbr As Integer
    
    
    curPos = 1
    ret = ""
    
    
    
    
    Dim i As Long
    Dim tokenBalance As Long
    
    i = 0
    tokenBalance = 0
    
    Do While i < 25 ' max of 25 tokens per string
        i = i + 1
    
        nextColNameStartPos = InStr(curPos, colTemplate, colDelimPrefix)
        
        If nextColNameStartPos = 0 Then
            If curPos <= Len(colTemplate) Then
                token = Right$(colTemplate, Len(colTemplate) - curPos + 1)
                'Debug.Print "- PartE: " & token
                
                If isFormula = False Then
                    token = """" & Replace(token, """", """""") & """"
                End If
                
                ret = ret & token
            Else
                If isFormula = False Then
                    ret = Left$(ret, Len(ret) - 3) ' Remove " & " at end
                End If
            End If
            
            Exit Do
        End If
        
        'If curPos > 1 Then
        '    If Mid$(str, curPos - 1, 1) = "\" Then
        '        curPos = curPos + Len(colDelimPrefix)
        '        ret = ret & colDelimPrefix
        '        skipIteration = True
        '    End If
        'End If
              
        If nextColNameStartPos >= 1 Then
            tokenBalance = tokenBalance + 1
            
            token = Mid$(colTemplate, curPos, nextColNameStartPos - curPos)
            'debug.Print "- PartM: " & token
            
            If isFormula = False Then
                token = """" & Replace(token, """", """""") & """ & "
            End If
            
            ret = ret & token
        End If
              

        nextColNameEndPos = InStr(curPos + Len(colDelimPrefix), colTemplate, colDelimSuffix)
    
        
        If nextColNameEndPos > 0 Then
            tokenBalance = tokenBalance - 1
            
            colName = Mid$(colTemplate, nextColNameStartPos + Len(colDelimPrefix), nextColNameEndPos - nextColNameStartPos - Len(colDelimPrefix))
            colNbr = Worksheet_GetColumnByName(sheet, colName)
            
            If colNbr = 0 Then Err.Raise -1, , "Cannot find column '" & colName & "' in worksheet " & sheet.Name
            
            ' 2017-08-22: bug fix: adjusts column number by one if new column is inserted before it
            If colNbr > insertAfterColumnNum Then colNbr = colNbr + 1
            
            If isFormula Then
                token = "R[0]C" & colNbr
            Else
                token = "R[0]C" & colNbr & " & "
            End If
            
            ret = ret & token
            
            'Debug.Print "- Col: " & colName
            
            curPos = nextColNameEndPos + Len(colDelimSuffix)
            
        End If
        
        
    Loop
    
    
    If tokenBalance <> 0 Then
        Err.Raise -1, , "Invalid template string for column '" & newColName & "'. Unbalanced column name parenthesis." & vbCrLf & vbCrLf & colTemplate
    End If
    
    
    If isFormula Then
        If Left(ret, 1) <> "=" Then ret = "=" & ret
    Else
        ret = "=" & ret
    End If
    
    Dim rowCount As Long, col As Long
    Dim rng As Range
    
    rowCount = Worksheet_GetLastRow(sheet)

    col = insertAfterColumnNum + 1
    sheet.columns(col).Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    sheet.Cells(1, col) = newColName
    
    If isFormula And isArrayFormula Then
        For i = 2 To rowCount
            sheet.Range(sheet.Cells(i, col), sheet.Cells(i, col)).FormulaArray = ret
        Next i
    Else
        sheet.Range(sheet.Cells(2, col), sheet.Cells(rowCount, col)).FormulaR1C1 = ret
    End If

    'rng.FormulaR1C1 = "=" & ret


End Sub

Public Sub Worksheet_Append(wsDest As Worksheet, wsSrc As Worksheet, Optional removeSrcHeaderRow As Boolean = True)
             
    ' Appends data from one worksheet to another, Useful for combining multiple sheets into one, provided the columns are the same
    
             
    Dim srcLastRow As Long, srcLastCol As Long
    Dim destLastRow As Long
    
    
    destLastRow = Worksheet_GetLastRow(wsDest)
    
    srcLastRow = Worksheet_GetLastRow(wsSrc)
    srcLastCol = Worksheet_GetLastColumn(wsSrc)
    
    ' Begin - Append data in temporary sheet to final worksheet
    Dim destRng As Range
    Dim srcRng As Range
    
        
    If destLastRow = 0 Or removeSrcHeaderRow = False Then
        ' Include header columns
        Set srcRng = wsSrc.Cells(1, 1).Resize(srcLastRow, srcLastCol)
        Set destRng = wsDest.Cells(1, 1).Resize(srcLastRow, srcLastCol)
    Else
        ' Do not include header columns
        Set srcRng = wsSrc.Cells(1, 1).Offset(1).Resize(srcLastRow - 1, srcLastCol)
        Set destRng = wsDest.Cells(1, 1).Offset(destLastRow).Resize(srcLastRow - 1, srcLastCol)
    End If

    Dim data() As Variant
    
    ' fast method
    data = srcRng
    destRng = data
    
    
    
End Sub
Public Sub Worksheet_AppendColumns(wsDest As Worksheet, wsSrc As Worksheet, columns As Variant)

    ' Appends data from one worksheet to another, using only data from specific columns
    
    
    'Dim wsDest As Worksheet: Set wsDest = Sheets("TEST")
    'Dim wsSrc As Worksheet: Set wsSrc = Sheets("ALPT")
    'Dim columns As Variant: columns = Array("DAY", "HR", "SITE", "ENODEB", "EUTRANCELL", "CARRIER")
    

    Dim srcLastRow As Long, srcLastCol As Long
    Dim destLastRow As Long
    
    Dim data() As Variant
    Dim destRng As Range
    
    
    
    destLastRow = Worksheet_GetLastRow(wsDest)
    
    
    data = Data_LoadFromRange(wsSrc.Cells, columns, IIf(destLastRow = 0, True, False), True)
    
    On Error Resume Next
    srcLastRow = UBound(data, 1) - LBound(data, 1) + 1 'Worksheet_GetLastRow(wsSrc)
    srcLastCol = UBound(data, 2) - LBound(data, 2) + 1
    On Error GoTo 0
    
    
    
    If srcLastRow > 0 And srcLastCol > 0 Then
        Set destRng = wsDest.Cells(destLastRow + 1, 1).Resize(srcLastRow, srcLastCol)
        
        ' fast method
        destRng = data
    End If
    

End Sub

Public Sub Workbook_SaveCopyWithoutMacros(destFile As String, Optional destFileFormat As XlFileFormat = XlFileFormat.xlOpenXMLWorkbook, Optional wb As Workbook = Nothing)

    ' TODO: Update to work with XLSB and file formats which support macros

    Dim destWB As Workbook
    Dim tmpFile As String, tmpFileSuffix As String
    Dim fileExt As String, fileExtPos As Long
    
    'Dim finalFileExt As String, finalFileFormat As XlFileFormat
    
    
    If wb Is Nothing Then Set wb = ThisWorkbook
    
    ' default:
    'fileExtFinal = "xls"
    'finalFileFormat


    fileExt = ""
    
    fileExtPos = InStrRev(wb.Name, ".")
    If fileExtPos = 0 And Len(wb.Name) - fileExtPos <= 4 Then ' Assume file extensions cannot be greater than 4 characters
        fileExt = Right$(fileExtPos, Len(wb.Name) - fileExtPos)
        tmpFileSuffix = "." & fileExt
    Else
        ' Deal with case where wb.Name does not have extension
        Select Case wb.FileFormat
            Case XlFileFormat.xlOpenXMLTemplate: tmpFileSuffix = ".xlt"
            Case XlFileFormat.xlOpenXMLTemplateMacroEnabled: tmpFileSuffix = ".xltm"
            Case XlFileFormat.xlOpenXMLWorkbook: tmpFileSuffix = ".xlsx"
            Case XlFileFormat.xlOpenXMLWorkbookMacroEnabled: tmpFileSuffix = ".xlsm"
            Case XlFileFormat.xlExcel12: tmpFileSuffix = ".xlsb"
            Case XlFileFormat.xlExcel8: tmpFileSuffix = ".xls"
            Case Else: tmpFileSuffix = ".xls"
        End Select
    End If
    
    
    ' Save copy as temporary file
    tmpFile = CreateTempFileName(wb.Name) & tmpFileSuffix
    wb.SaveCopyAs tmpFile
    

    ' Open temp file first re-save as OpenWorkbook format (without macros), and then re-save as destination format
    Application.DisplayAlerts = False
    
    ' re-save as OpenWorkbook format (without macros)
    Set destWB = Workbooks.Open(tmpFile)
    destWB.CheckCompatibility = False
    destWB.SaveAs fileName:=tmpFile & ".xlsx", FileFormat:=XlFileFormat.xlOpenXMLWorkbook, CreateBackup:=True, AccessMode:=XlSaveAsAccessMode.xlExclusive
    destWB.SaveAs fileName:=destFile, FileFormat:=destFileFormat, CreateBackup:=True, AccessMode:=XlSaveAsAccessMode.xlExclusive                ' Save to destination in file format
    destWB.Close False
    Set destWB = Nothing
    
    ' Open workbook without macros and save to final destination
    Set destWB = Workbooks.Open(tmpFile & ".xlsx")
    destWB.CheckCompatibility = False
    destWB.SaveAs fileName:=destFile, FileFormat:=destFileFormat, CreateBackup:=True, AccessMode:=XlSaveAsAccessMode.xlExclusive
    destWB.Close False
    Set destWB = Nothing
    
    Application.DisplayAlerts = True
    
End Sub

' Workbook_RefreshPivotTables: Refresh pivot tables matching names. The names may contain wildcards
'
' Examples:
'   1) Refresh pivot tables named PivotTable1, PivotTable2
'       Workbook_RefreshPivotTables thisworkbook, "PivotTable1", "PivotTable2"
'   2) Refresh all pivot tables which start with "PivotTable" (PivotTable1, PivotTable2, PivotTable3, ...)
'       Workbook_RefreshPivotTables thisworkbook, "PivotTable1", "PivotTable2"

Public Sub Workbook_RefreshPivotTables(wb As Workbook, ParamArray pivotTableNames() As Variant)

    Dim ws As Worksheet, pt As PivotTable, ptNameVar As Variant, ptNameStr
    
    For Each ws In wb.Worksheets
        For Each pt In ws.PivotTables
        
            For Each ptNameVar In pivotTableNames
                ptNameStr = CStr(ptNameVar)
                
                If pt.Name Like ptNameStr Then
                    pt.RefreshTable
                End If
            Next ptNameVar
            
        Next pt
    Next ws
    

End Sub



Public Function Array_Count(ByRef arr As Variant) As Long

    Array_Count = 0
    
On Error GoTo ItemNotOrArrayErr:

    If Not IsArray(arr) Then Exit Function
    If IsEmpty(arr) Then Exit Function
    
    Array_Count = UBound(arr) - LBound(arr) + 1
    Exit Function
    
ItemNotOrArrayErr:
    

End Function
Public Function Array_IsAllocated(arr As Variant) As Boolean
    
    On Error Resume Next
    Array_IsAllocated = IsArray(arr) And _
                        Not IsError(LBound(arr, 1)) And _
                        LBound(arr, 1) <= UBound(arr, 1)
                        
End Function
                           
Public Function Array_Append(ByVal arr As Variant, ByVal append As Variant) As Variant


    Dim origLB As Long, origUB As Long
    Dim numItems As Long
    
    
    If Not IsArray(arr) Then
        If IsEmpty(arr) Then
            Array_Append = Array(append)
        Else
            Array_Append = Array(arr, append)
        End If
        
        Exit Function
    End If
    
    
    On Error GoTo ArrayNotDimensioned
    
    origLB = LBound(arr)
    origUB = UBound(arr)
    
    On Error GoTo 0
    
    
    ReDim Preserve arr(origLB To origUB + 1) As Variant
    
    arr(origUB + 1) = append

    Array_Append = arr
    Exit Function
    
ArrayNotDimensioned:
    ReDim arr(1 To 1) As Variant
    arr(1) = append
    Array_Append = arr
    Exit Function

End Function

Function Array_AppendMultiple(ByVal arr As Variant, ByVal append As Variant) As Variant


    Dim origLB As Long, origUB As Long
    Dim numItems As Long
    Dim i As Integer, j As Integer
    
    If Not IsArray(append) Then append = Array(append)
    
    If Not IsArray(arr) Then
        If IsEmpty(arr) Then
            arr = Array()
        Else
            arr = Array(arr)
        End If
    End If
    
    
    numItems = 0
    origLB = 0
    origUB = -1
    
    On Error Resume Next ' Trap errors - keeps numItems at 0 if error occurs
    origLB = LBound(arr)
    origUB = UBound(arr)
    numItems = UBound(append) - LBound(append) + 1
    On Error GoTo 0
    
    If numItems > 0 Then
        ReDim Preserve arr(origLB To origUB + numItems) As Variant
        
        For i = LBound(append) To UBound(append)
            arr(origUB + i - LBound(append) + 1) = append(i)
        Next i
    End If
    
    Array_AppendMultiple = arr

End Function
Function Array_CreateAndFill(numItems As Long, defaultItem As Variant) As Variant

    If numItems = 0 Then
        Array_CreateAndFill = Array()
        Exit Function
    End If

    Dim retArr() As Variant
    Dim i As Integer
    
    ReDim retArr(1 To numItems) As Variant
    
    For i = LBound(retArr) To UBound(retArr)
        retArr(i) = defaultItem
    Next i
    
    Array_CreateAndFill = retArr

End Function
Function Array_Find(arr As Variant, search As Variant, Optional occurance As Integer = 1, Optional notFoundValue As Long = -1) As Long

On Error GoTo ErrOccurred:

    Dim i As Long
    Dim curOcc As Long
    
    Array_Find = notFoundValue
    
    If Not IsArray(arr) Then Exit Function
    If UBound(arr) = -1 Then Exit Function
    
    
    curOcc = 0
    
    For i = LBound(arr) To UBound(arr)
        If arr(i) = search Then
            curOcc = curOcc + 1
            
            If curOcc = occurance Then
                Array_Find = i
                Exit Function
            End If
        End If
    Next i
    
    
ErrOccurred:

End Function

Function Array_ItemExists(arr As Variant, search As Variant) As Boolean

    Array_ItemExists = (Array_Find(arr, search) >= 0)

End Function
Function Array_ToString(arr As Variant, Optional formatStr As String) As String()

    Dim i As Long
    Dim cnt As Long
    Dim ret() As String
    Dim arrItem As Variant
    
    cnt = Array_Count(arr)
    
    If cnt > 0 Then
        ReDim ret(LBound(arr) To UBound(arr)) As String
        
        For i = 0 To cnt - 1
            arrItem = arr(i + LBound(arr))
            
            If Len(formatStr) > 0 Then
                ret(i + LBound(arr)) = Format$(arrItem, formatStr)
            Else
                ret(i + LBound(arr)) = CStr(arrItem)
            End If
        Next i
    End If

    Array_ToString = ret

End Function
Public Function Array_Combine(ParamArray arrs() As Variant) As Variant

    Dim numArrays As Long, arrayCnt As Long, testCnt As Long
    Dim i As Integer, j As Integer
    
    numArrays = UBound(arrs) - LBound(arrs) + 1
    arrayCnt = 0
    
    For i = LBound(arrs) To UBound(arrs)
        testCnt = -1
        If IsArray(arrs(i)) Then testCnt = Array_Count(arrs(i))
        
        If arrayCnt > 0 Then
            If testCnt <> arrayCnt Then Err.Raise -1, , "Array_Merge: Number of items in each dimension must be equal"
        Else
            arrayCnt = testCnt
        End If
    Next i
    
    
    Dim retArray() As Variant
    ReDim retArray(1 To numArrays, 1 To arrayCnt) As Variant
    
    For i = LBound(arrs) To UBound(arrs)
        For j = 1 To arrayCnt
            retArray(i - LBound(arrs) + 1, j) = arrs(i)(j + LBound(arrs(i)) - 1)
        Next j
    Next i
    
    Array_Combine = retArray

End Function
Function Array_NumDimensions(arr As Variant) As Long



    If Not IsArray(arr) Then
        Array_NumDimensions = 0
        Exit Function
    End If
      
    
    On Error GoTo FinalDimension
    
    Dim i As Long, errCheck As Long
    
    'VBA arrays can have up to 60000 dimensions; this allows for that.
    For i = 1 To 60000
        'It is necessary to do something with the LBound to force it
        'to generate an error.
        errCheck = LBound(arr, i)
    Next i
      

FinalDimension:

      Array_NumDimensions = i - 1

End Function
Public Function ConCat(delimiter As Variant, ParamArray cellRanges() As Variant) As String

    Dim returnStr As String
    Dim cell As Range, Area As Variant
    Dim ArrayItem As Variant

    If IsMissing(delimiter) Then delimiter = ""

    For Each Area In cellRanges
        Dim typeN As Variant
        typeN = TypeName(Area)
    
        If Not IsError(Area) Then
            If TypeName(Area) = "Range" Then
                For Each cell In Area
                    If Not IsError(cell) Then
                        If Len(cell.Value) Then returnStr = returnStr & delimiter & cell.Value
                    End If
                Next
            ElseIf IsArray(Area) Then
                For Each ArrayItem In Area
                    If Not IsError(ArrayItem) Then
                        If Len(ArrayItem) Then returnStr = returnStr & delimiter & ArrayItem
                    End If
                Next
            Else
                returnStr = returnStr & delimiter & Area
            End If
        End If
    Next

    ConCat = Mid(returnStr, Len(delimiter) + 1)
    
End Function
' Resizes 1-D to a 2D array with a specific row/col count. Overflow occurs when the number of items is greater than rows*col
Public Function Array_ResizeOverflow(arr() As Variant, rows As Long, cols As Long, Optional delim As String = "|") As Variant

    ' joe h:    Added 5/31/16

    'Dim arr() As Variant: arr = Array(1, 2, 3, 4, 5, 6, 8, 9)
    'Dim rows As Long: rows = 2
    'Dim cols As Long: cols = 3
    'Dim delim As String: delim = "|"
    'Dim overflowMode As Integer: overflowMode = 1
    


    ' Overflow modes ( Work in Progress)
    '  - 0: OverflowOnLastElement
    '  - 1: OverflowOnRows
    '  - 2: OverflowOnCols

    Dim inOverflowMode As Boolean
    Dim arrIdx As Long
    Dim newArr() As Variant
    
    Dim mapRow As Long, mapCol As Long
    
    ReDim newArr(1 To rows, 1 To cols) As Variant
    
    For arrIdx = LBound(arr) To UBound(arr)
        mapRow = Int((arrIdx - LBound(arr)) / cols) + 1
        mapCol = ((arrIdx - LBound(arr)) Mod cols) + 1
        
        'Start overflowing when we are past the number of entries
        inOverflowMode = (arrIdx - LBound(arr)) > rows * cols - 1
        
        If inOverflowMode Then
            'Select Case overflowMode
            '    Case 1: ' Overflow On Rows
            '        ' Map to last element in same row
            '        mapCol = cols
            '    Case Else: ' Overflow on LastElement -> Start overflowing when there are no more entries in mapRow
            '        ' Map to last element
            '        mapRow = rows
            '        mapCol = cols
            'End Select
            
            mapRow = rows
            mapCol = cols
            
            newArr(mapRow, mapCol) = newArr(mapRow, mapCol) & delim & arr(arrIdx)
        Else
            newArr(mapRow, mapCol) = arr(arrIdx)
        End If

        
        
        'Debug.Print arr(arrIdx), mapRow, mapCol
    Next arrIdx


    'Dim i As Long, j As Long
    'Dim str As String
   '
   ' For i = 1 To rows
   '     str = ""
   '     For j = 1 To cols
   '         str = str & newArr(i, j) & vbTab
   '     Next j
   '     Debug.Print str
   ' Next i

    Array_ResizeOverflow = newArr

End Function

' Acts like V/HLookup but returns all matches as a delimited list
Function Lookup_MultiMatch(lookupValue As Variant, lookupRange As Range, resultsRange As Range, Optional delim As String = ", ") As String

    ' 2015-05-13 - Added
    

    Dim s As String 'Results placeholder
    Dim sTmp As String  'Cell value placeholder
    Dim r As Long   'Row
    Dim c As Long   'Column
    Const tmpDelimiter = "|||"  'Makes InStr more robust

    s = tmpDelimiter
    For r = 1 To lookupRange.rows.count
        For c = 1 To lookupRange.columns.count
            If lookupRange.Cells(r, c).Value = lookupValue Then
                'I know it's weird to use offset but it works even if the two ranges
                'are of different sizes and it's the same way that SUMIF works
                sTmp = resultsRange.Offset(r - 1, c - 1).Cells(1, 1).Value
                If InStr(1, s, tmpDelimiter & sTmp & tmpDelimiter) = 0 Then
                    s = s & sTmp & tmpDelimiter
                End If
            End If
        Next
    Next


    s = Replace(s, tmpDelimiter, delim)
    If Left(s, 1) = "," Then s = Mid(s, 2)
    If Right(s, 1) = "," Then s = Left(s, Len(s) - 1)

    Lookup_MultiMatch = s
    
End Function


Function FindN(sFindWhat As String, sInputString As String, n As Integer) As Integer
    
    Dim j As Integer
    
    FindN = 0
    
    For j = 1 To n
        FindN = InStr(FindN + 1, sInputString, sFindWhat)
        If FindN = 0 Then Exit For
    Next j
    
End Function
Public Function Token_Extract_WIP(str As String, delim As String, n As Integer) As String
        
     
    ' added 10/28
    
    
    Dim pos1 As Long, pos2 As Long
    Dim i As Integer
    
    pos1 = 0
    
    For i = 1 To n
        pos1 = pos2
        pos2 = InStr(pos1 + Len(delim), str, delim)
        
        If pos2 = 0 Then
            If i < n Then
                Debug.Print "no token in position" & n
                Token_Extract_WIP = vbNullString
                Exit Function
            Else
                pos2 = Len(str)
            End If
            
            Exit For
        End If
        
        If i = n Then Exit For
        
    Next i

    Token_Extract_WIP = Mid$(str, pos1 + Len(delim), pos2 - pos1 - Len(delim))
    
    Debug.Print
    
End Function
' Gets Range object from text, worksheet, or range itself
Public Function Range_FromAny(table As Variant) As Range

    ' 2016-06-03 -  renamed from GetTableRange()
    
    Dim ws As Worksheet
    Dim outRng As Range
    Dim lastCol As Long, lastRow As Long
    
    Set outRng = Nothing
    
    Select Case TypeName(table)
        Case "String":
            'Assume table is the name of a worksheet
            
            If Worksheet_SheetExists(CStr(table)) Then
                Set ws = ThisWorkbook.Sheets(table)
                
                lastRow = Worksheet_GetLastRow(ws)
                lastCol = Worksheet_GetLastColumn(ws)
                
                Set outRng = ws.Cells(1, 1).Resize(lastRow, lastCol)
            Else
                Err.Raise -1, , "Sheet '" & table & "' does not exist"
            End If

        
        Case "Worksheet"
            Set ws = ThisWorkbook.Sheets(table)
            
            lastRow = Worksheet_GetLastRow(ws)
            lastCol = Worksheet_GetLastColumn(ws)
            
            Set outRng = ws.Cells(1, 1).Resize(lastRow, lastCol)
           
        Case "Range"
            Set outRng = table
     
        Case Else
            Err.Raise -1, , "Unrecognized data source: " & table
    End Select
    
    
    Set Range_FromAny = outRng


End Function
Function Range_GetColumnByName(rng As Range, colName As String, Optional headerRow As Long = 1) As Long
    
    ' 2016-06-10 - Changed from Range_GetColumnNum() to Range_GetColumnByName()
    ' 2015-05-13 - Added headerRow param


    Dim r As Range
    
    Set r = rng.rows(headerRow).Find(colName, LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlNext, matchCase:=True)
    
    If r Is Nothing Then
        Range_GetColumnByName = 0
    Else
        Range_GetColumnByName = r.column
    End If
    
End Function
' Resolves columns names within an array to its column numbers. Existing numbers are left alone.
Public Function Range_GetColumnsByName(rng As Range, columns() As Variant, Optional headerRow As Long = 1) As Long()

    Dim ret() As Long
    Dim i As Long
    
    Dim tableName As String
    tableName = rng.ListObject.Name
    'Set ActiveTable = ActiveSheet.ListObjects(tableName)
    
    ReDim ret(LBound(columns) To UBound(columns)) As Long

    For i = LBound(columns) To UBound(columns)
        If IsNumeric(columns(i)) Then
            ret(i) = columns(i)
        Else
            ret(i) = Range_GetColumnByName(rng, CStr(columns(i)), headerRow)
        End If
    Next i
    
    Range_GetColumnsByName = ret

End Function
' Recursively resolves columns names within an array to its column numbers. Existing numbers are left alone. Recursive
Public Function Range_GetColumnsByNameRecursive(rng As Range, columns As Variant, Optional headerRow As Long = 1) As Variant ' Recursive

    Dim ret As Variant
    Dim i As Long
    
    If IsArray(columns) Then
        ret = columns
        
        For i = LBound(columns) To UBound(columns)
            ret(i) = Range_GetColumnsByNameRecursive(rng, columns(i), headerRow)
        Next i
    Else
        If IsNumeric(columns) Then
            ret = columns
        Else
            ret = Range_GetColumnByName(rng, CStr(columns), headerRow)
        End If
    End If
    
    Range_GetColumnsByNameRecursive = ret

End Function
Public Function Range_LastRow(rng As Range) As Long
    On Error Resume Next
    Range_LastRow = rng.Cells.Find(What:="*", After:=rng.Cells(1, 1), LookAt:=xlPart, LookIn:=xlFormulas, SearchOrder:=xlByRows, SearchDirection:=xlPrevious, matchCase:=False).row - rng.Cells(1, 1).row + 1
    On Error GoTo 0
End Function
Public Function Range_LastColumn(rng As Range) As Long
    On Error Resume Next
    Range_LastColumn = rng.Cells.Find(What:="*", After:=rng.Cells(1, 1), LookAt:=xlPart, LookIn:=xlFormulas, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, matchCase:=False).row - rng.Cells(1, 1).column + 1
    On Error GoTo 0
End Function
' Range_LoadColumnDataToArray: Loads specific columns from table within a range
' DataTable_GetByName: returns the first ListObject found in the workbook with the provided name
Public Function DataTable_GetByName(tblName As String, Optional wb As Workbook = Nothing) As ListObject
    
    
    If wb Is Nothing Then: Set wb = ThisWorkbook
    
    Dim ws As Worksheet, listObj As ListObject
    
    For Each ws In wb.Sheets
        For Each listObj In ws.ListObjects
            If listObj.Name = tblName Then
                Set DataTable_GetByName = listObj
                Exit Function
            End If
        Next listObj
    Next ws
    
    
    Set DataTable_GetByName = Nothing

End Function
' DataTable_InsertBlankRows: Fast way to add many blank rows to table without using ListObject.ListRows.Add() - 100x speed increase
Public Function DataTable_InsertBlankRows(loTable As ListObject, numRows As Long) As Range

    Dim prevShowTotals As Boolean
    Dim rng As Range, firstColRng As Range
    
    prevShowTotals = loTable.ShowTotals
    loTable.ShowTotals = False

    
    If loTable.InsertRowRange Is Nothing Then
        Set rng = loTable.ListRows.Add.Range
    Else
        Set rng = loTable.InsertRowRange
    End If
    
    
    'Application.ScreenUpdating = False
    Excel_AppUpdates_Disable
    
    Set firstColRng = rng.Resize(numRows, 1)
    firstColRng.Value = Chr(160)        ' Set to non-breaking space
    firstColRng.Value = vbNullString    ' delete non-breaking space
    
    Excel_AppUpdates_Restore
    'Application.ScreenUpdating = True
    
    
    loTable.ShowTotals = prevShowTotals
    
    Set DataTable_InsertBlankRows = rng.Resize(numRows, loTable.ListColumns.count)
    

End Function
' DataTable_AppendWorksheet: Appends data from a worksheet into a datatable
Public Sub DataTable_AppendWorksheet(loTable As ListObject, ws As Worksheet, Optional wsHeaderRow As Long = 1, Optional customMapDict As Dictionary = Nothing)

    
    Dim i As Long
    
    Dim wsColumnDict As Dictionary   ' Scripting.Dictionary
    
    Dim numRowsToAdd As Long
    Dim numTableCols As Long, numTableRows As Long
    
    Dim tblColumns() As String     ' Column name
    Dim tblColumnMapping() As Long ' Holds index of each mapped column. Value of zero (0) indicates no mapping
    
    Dim lstCol As ListColumn, searchColName As String
    Dim tblColIndex As Long
    
    
    numRowsToAdd = Worksheet_GetLastRow(ws, wsHeaderRow) - wsHeaderRow
    
    numTableCols = loTable.ListColumns.count
    numTableRows = loTable.ListRows.count
    
    Set wsColumnDict = Worksheet_GetColumnDictionary(ws)
    
    
    
    ReDim tblColumns(1 To numTableCols) As String
    ReDim tblColumnMapping(1 To numTableCols) As Long
    
    ' Map existing table columns to worksheet's columns (if a column with same name exists or mapped explicitly using customMapDict)
    For Each lstCol In loTable.ListColumns
        ' TODO: add a check here to allow user to restrict columns to copy by some method
        
        tblColumns(lstCol.Index) = lstCol.Name
        
        searchColName = lstCol.Name
        
        If Not customMapDict Is Nothing Then
            If customMapDict.Exists(lstCol.Name) Then searchColName = customMapDict.Item(lstCol.Name)
        End If
     
        
        ' Search source worksheet for corresponding column and map it
        If wsColumnDict.Exists(searchColName) Then
            tblColumnMapping(lstCol.Index) = wsColumnDict.Item(searchColName)
        End If
        
        
    Next lstCol
    
    Set wsColumnDict = Nothing ' free memory
    
    
    ' TODO: Optional: Add checks to verify mappings are good (need to define 'good')
    
    
    Dim rngInsertedRows As Range
    Set rngInsertedRows = DataTable_InsertBlankRows(loTable, numRowsToAdd)
    
    
    Excel_AppUpdates_Disable
    
    
    ' Loop through each mapped column and copy from source worksheet column to destination table column
    Dim rngDest As Range, data() As Variant
    Dim rngSrc As Range
    
    
    For tblColIndex = 1 To numTableCols
        If tblColumnMapping(tblColIndex) Then
            Set rngDest = loTable.DataBodyRange.Offset(numTableRows, tblColIndex - 1).Resize(numRowsToAdd, 1)
            Set rngSrc = ws.Cells(wsHeaderRow + 1, tblColumnMapping(tblColIndex)).Resize(numRowsToAdd, 1)
            
            data = rngSrc
            rngDest = data
        End If
    Next tblColIndex

    Excel_AppUpdates_Restore




End Sub

Public Function Collection_ToArray(c As Collection) As Variant()
    Dim a() As Variant: ReDim a(0 To c.count - 1)
    Dim i As Long
    For i = 1 To c.count
        a(i - 1) = c.Item(i)
    Next
    Collection_ToArray = a
End Function
Public Function Collection_ToArrayLong(c As Collection) As Long()
    Dim a() As Long: ReDim a(0 To c.count - 1) As Long
    Dim i As Long
    For i = 1 To c.count
        a(i - 1) = c.Item(i)
    Next
    Collection_ToArrayLong = a
End Function
' Calculates the average value from set of data points within the range of a given percentile
Public Function PercentileAverage(dataPoints() As Double, Optional pcntileBoundStart As Double = 0, Optional pcntileBoundEnd As Double = 1) As Double


    Debug.Assert pcntileBoundStart >= 0 And pcntileBoundStart <= 1
    Debug.Assert pcntileBoundEnd >= 0 And pcntileBoundEnd <= 1
    Debug.Assert pcntileBoundStart <= pcntileBoundEnd
    

    Dim i As Long, j As Long
    Dim rankStart As Integer, rankEnd As Integer
    Dim numDataPoints As Long
    
    
    numDataPoints = Array_Count(dataPoints)
    
    
    ' default: include all
    rankStart = 1
    rankEnd = numDataPoints
    
    
    ' sort data and calculate start/end ranks for given percentiles
    If pcntileBoundStart > 0 Or pcntileBoundEnd < 1 Then
        ' Sort ascending
        Dim tmpDP As Double
        
        For i = LBound(dataPoints) To UBound(dataPoints)
            For j = i + 1 To UBound(dataPoints)
                If dataPoints(j) < dataPoints(i) Then
                    tmpDP = dataPoints(i)
                    dataPoints(i) = dataPoints(j)
                    dataPoints(j) = tmpDP
                End If
            Next j
        Next i
    
        rankStart = Round(pcntileBoundStart * (numDataPoints - 1) - 0.5) + 1
        rankEnd = Round(pcntileBoundEnd * (numDataPoints - 1) + 0.5) + 1
    
        If rankStart < 1 Then rankStart = 1
        If rankEnd > numDataPoints Then rankEnd = numDataPoints
    End If
    
    
    Dim avg As Double, count As Long
    
    avg = 0
    count = 0
    
    ' Calculate average between ranks (numerically stable algorithm)
    For i = rankStart To rankEnd
        count = count + 1
        avg = avg + ((dataPoints(i - 1 + LBound(dataPoints)) - avg) / count)
    Next i
    
    
    PercentileAverage = avg


End Function
' Table_GroupBy_Map: Takes a table within spreadsheet, groups the rows, and returns a mapping for additional processing.
'    The mapping can then be iterated through to obtain row data or perform specific calculations for a particular group.
'
'    TableGroupByMapping has two data structures: GroupToRowMap, RowToGroupMap which provide the necessary values needed to
'    iterate over the data
'
'    GroupToRowMap() is an nested array of arrays, with <number of group columns> + 1 levels. The last level contains the row numbers
'        associated with the specific unique combination of the group values.
'
'        The nessted array GroupToRowMap(G_1)(G_2)...(G_N) contains an array where the elements 1..M contains the row numbers for each
'        unique group pair
'
'
'    RowToGroupMap
'
'    Example Table:
'            ------------------------------------
'    Row 1:  | Name   | Title        | Manager  |
'            ------------------------------------
'    Row 2:    John       Engineer       Chris
'    Row 3:    Casey      Specialist     Chris
'    Row 4:    Dan        Engineer       Mike
'    Row 5:    Bob        Engineer       Mike
'
'    Example 1a:
'        Table_GroupBy_Map(<ws>, Array("Title")) will group by Title.
'          In this case, 3 rows will belong to the group value Engineer, and 1 row to the group Specialist
'
'          TableGroupByMapping.GroupToRowMap will return a nested array of arrays 2 levels deep
'          Format: <array index>: <value>
'          1:
'               1.0: "Engineer"
'               1.1: 2 <row number of 1st Engineer>
'               1.2: 4 <row number of 2nd Engineer>
'               1.2: 5 <row number of 3rd Engineer>
'          2:
'               2.0: "Specialist"
'               2.1: 3 <row number of 1st Specialist>
'
'    Example 2a:
'        Table_GroupBy_Map(<ws>, Array("Manager", "Title")) will group by Manager, then Title.
'          In this case, 3 unique groups will be found in the data:
'            Group Chris will have 2 subgroups (Engineer,Specialist) and each will be associated with 1 row
'            Group Mike will have 1 subgroup (Engineer) and will be associated with 2 rows
'
'          TableGroupByMapping.GroupToRowMap will return a nested array of arrays 3 levels deep
'          Format: <array index>: <value>
'          1:
'               1.0: "Chris"
'               1.1: Array
'                   1.1.0: "Engineer"
'                   1.1.1: 2 <row number of 1st Engineer managed by Chris>
'               1.2: Array
'                   1.1.0: "Specialist"
'                   1.1.1: 3 <row number of 1st Specialist managed by Chris>
'          2:
'               2.0: "Mike"
'               2.1: Array
'                   2.1.0: "Engineer"
'                   2.1.1: 2 <row number of 1st Engineer managed by Mike>
'                   2.1.2: 3 <row number of 2nd Engineer managed by Mike>
'
Public Function Table_GroupBy_Map(ws As Worksheet, groupColumns() As Variant, Optional headerRow As Long = 1, Optional groupToRowMapIncludeValueAsZeroIndexItem As Boolean = False) As TableGroupByMapping



    Dim gMap As TableGroupByMapping

    Dim groupColumnNums As Variant ' Array of Longs() or array of array of longs(),
    Dim groupColumnCount As Long

    groupColumnNums = Worksheet_GetColumnsByName(ws, groupColumns, headerRow)
    groupColumnCount = Array_Count(groupColumns)
    
    
    Dim lastRow As Long, curRow As Long
    Dim pCol As Long
    Dim i As Long, j As Long, k As Long
    
    lastRow = Worksheet_GetLastRow(ws)
    
    Dim curGroupColumn As Variant
    Dim curGroupColumnID As String
    
    
    Dim rowPartialIDs() As String
    ReDim rowPartialIDs(1 To groupColumnCount) As String
    Dim rowGroupColumnIDs() As String
    ReDim rowGroupColumnIDs(1 To groupColumnCount) As String
    
    
    ReDim gMap.RowToGroupMap(headerRow + 1 To lastRow, 1 To groupColumnCount + 1) As Long
    
    
    Const UPV_MULTI_COL_SEP As String = "||||"
    Const UPV_PARENT_CHILD_COL_SEP As String = "@@@@"
    
    
    For curRow = headerRow + 1 To lastRow
        ' Calculate the ID for each group columns
        For pCol = LBound(groupColumnNums) To UBound(groupColumnNums)
            ' Group column is made up of two or more worksheet columns
            If IsArray(groupColumnNums(pCol)) Then
                curGroupColumnID = ""
                
                For j = LBound(groupColumnNums(pCol)) To UBound(groupColumnNums(pCol))
                    If IsArray(groupColumnNums(pCol)(j)) Then Err.Raise -1, , "Group columns may only be a single column (scalar) or multi column (array of scalars)"
                    
                    curGroupColumnID = curGroupColumnID & UPV_MULTI_COL_SEP & ws.Cells(curRow, groupColumnNums(pCol)(j))
                Next j
                
                curGroupColumnID = Mid(curGroupColumnID, UPV_MULTI_COL_SEP + 1)
            Else
                curGroupColumnID = ws.Cells(curRow, groupColumnNums(pCol))
            End If
            
            rowGroupColumnIDs(pCol - LBound(groupColumnNums) + 1) = curGroupColumnID
            
            If pCol = LBound(groupColumnNums) Then
                rowPartialIDs(pCol - LBound(groupColumnNums) + 1) = curGroupColumnID
            Else
                rowPartialIDs(pCol - LBound(groupColumnNums) + 1) = rowPartialIDs(pCol - LBound(groupColumnNums)) & UPV_PARENT_CHILD_COL_SEP & curGroupColumnID
            End If
            
            'Debug.Print curGroupColumnID
        Next pCol
        
        Dim idx As Long, parentIdx As Long, relIdx As Long
        
        'Dim groupColumnCurrentIndex() As Long
        'ReDim groupColumnCurrentIndex(1 To groupColumnCount) As Long
        
        
        Dim rowHasNewPartialRowID As Boolean
        
        
        Dim rootChildCount As Long
        
        Dim uniqueGroupValues() As TableGroupByMap_UniqueGroupValue
        Dim uniqueGroupValue_Count As Long
        
        ' For the current row, rowGroupColToUPV() stores the index of the uniquePartialRow entry representing the respective group column
        Dim rowGroupColToUPV() As Long
        ReDim rowGroupColToUPV(1 To groupColumnCount) As Long
        
        '
        rowHasNewPartialRowID = False
        
        ' Reset
        For pCol = 1 To groupColumnCount: rowGroupColToUPV(pCol) = 0: Next pCol
        
        For pCol = 1 To groupColumnCount
            idx = 0
            
            ' Begin - Find existing UPV
            
            ' if a previous group column in the same row created a new entry in uniqueGroupValues(), then all remaining group columns
            ' will also create a new entry. In this case, do not bother searching (rowHasNewPartialRowID = False)
            If rowHasNewPartialRowID = False Then
                'We will restrict our search to unique IDs who has (1) same group col (2) the same parent ID
                parentIdx = 0
                If pCol > 1 Then parentIdx = rowGroupColToUPV(pCol - 1)
            
                For i = 1 To uniqueGroupValue_Count
                    If uniqueGroupValues(i).GroupCol = pCol And uniqueGroupValues(i).ParentIndex = parentIdx Then
                        If uniqueGroupValues(i).ID = rowPartialIDs(pCol) Then
                            idx = i
                            Exit For
                        End If
                    End If
                Next
            End If
            ' End - Find existing UPV
            
            ' No existing UPV -> Create one
            If idx = 0 Then
                rowHasNewPartialRowID = True
                
                uniqueGroupValue_Count = uniqueGroupValue_Count + 1
                idx = uniqueGroupValue_Count
                
                ReDim Preserve uniqueGroupValues(1 To uniqueGroupValue_Count) As TableGroupByMap_UniqueGroupValue
                
                
                ' Set parent index (previous column) - can be 0 for a root/1st level value
                parentIdx = 0
                If pCol > 1 Then parentIdx = rowGroupColToUPV(pCol - 1)
                
                
                ' Set initial UPV values
                uniqueGroupValues(idx).ID = rowPartialIDs(pCol)
                uniqueGroupValues(idx).Value = rowGroupColumnIDs(pCol)
                uniqueGroupValues(idx).GroupCol = pCol
                uniqueGroupValues(idx).ParentIndex = parentIdx
                uniqueGroupValues(idx).FirstRowInContext = curRow
                
                
                ' Since we added a new child to the parent, increment the parent child count and set the relative index based on the new count
                If parentIdx > 0 Then
                    uniqueGroupValues(parentIdx).ChildCount = uniqueGroupValues(parentIdx).ChildCount + 1
                    uniqueGroupValues(idx).RelativeIndex = uniqueGroupValues(parentIdx).ChildCount
                Else
                    ' Special case: root node
                    rootChildCount = rootChildCount + 1
                    uniqueGroupValues(idx).RelativeIndex = rootChildCount
                End If
            End If
            
            If pCol = groupColumnCount Then
                ' We're on the last column so children of this would be the rows themselves.
                
                uniqueGroupValues(idx).ChildCount = uniqueGroupValues(idx).ChildCount + 1
            End If
            
            rowGroupColToUPV(pCol) = idx
        
            gMap.RowToGroupMap(curRow, pCol) = uniqueGroupValues(idx).RelativeIndex
        Next pCol


        ' We now have corresponding entries in the uniqueGroupValues for this row's group values

        ' idx contains the last group column index entry ie. idx = rowGroupColToUPV(groupColumnCount)
        
        
        If uniqueGroupValues(idx).ChildCount = 1 Then Set uniqueGroupValues(idx).RowNumbers = New Collection
        uniqueGroupValues(idx).RowNumbers.Add curRow

        ' The last entry of the row to GroupMap is the relative index of each individual item within (e.g. the current child count)
        gMap.RowToGroupMap(curRow, groupColumnCount + 1) = uniqueGroupValues(idx).ChildCount
        
    Next curRow
    
    
    

    gMap.GroupToRowMap = Table_GroupBy_Map_BuildGroupToRowMap(0, uniqueGroupValues, groupToRowMapIncludeValueAsZeroIndexItem)
    

    Table_GroupBy_Map = gMap

End Function
' Table_GroupBy_Map_BuildGroupToRowMap: Recursive utility function to build GroupToRowMap mapping in the form of nested arrays with the
'   number of groups = the nest level. The last array contains the row numbers of all the items which belong to a particular group
Private Function Table_GroupBy_Map_BuildGroupToRowMap(parentIdx As Long, ByRef uniqueGroupValues() As TableGroupByMap_UniqueGroupValue, Optional includeValueAsZeroIndexItem = False) As Variant()


    'Dim includeValueAsZeroIndexItem As Boolean: includeValueAsZeroIndexItem = True
    'Dim includeCountAsZeroIndexItem As Boolean: includeCountAsZeroIndexItem = False

    Dim i As Long
    Dim count As Long
    Dim ret() As Variant, retIdx As Long
    Dim retStartIdx As Long
    
    retStartIdx = IIf(includeValueAsZeroIndexItem, 0, 1)
    
    count = 0
    
    ' First pass - get count
    For i = LBound(uniqueGroupValues) To UBound(uniqueGroupValues)
        If uniqueGroupValues(i).ParentIndex = parentIdx Then count = count + 1
    Next i
    
    
    If count = 0 Then ' We are at the leaf end and we need to return the the actual rows
        Debug.Assert parentIdx >= LBound(uniqueGroupValues)
        
        count = uniqueGroupValues(parentIdx).ChildCount
        
        ReDim ret(retStartIdx To count) As Variant
        
        
        'If includeCountAsZeroIndexItem Then ret(0) = count
        If includeValueAsZeroIndexItem And parentIdx >= LBound(uniqueGroupValues) Then ret(0) = uniqueGroupValues(parentIdx).ID
    
        'If parentIdx >= LBound(uniqueGroupValues) Then ret(0) = uniqueGroupValues(parentIdx).ChildCount
        
        Debug.Assert Not (uniqueGroupValues(parentIdx).RowNumbers Is Nothing)
        
        Dim col As Collection
        Set col = uniqueGroupValues(parentIdx).RowNumbers
        
        For i = 1 To col.count
            ret(i) = col(i)
        Next i
        
        Set col = Nothing

        Table_GroupBy_Map_BuildGroupToRowMap = ret
        Exit Function
    End If
    
    
    ReDim ret(retStartIdx To count) As Variant
    
    'If includeCountAsZeroIndexItem Then ret(0) = count
    If includeValueAsZeroIndexItem And parentIdx >= LBound(uniqueGroupValues) Then ret(0) = uniqueGroupValues(parentIdx).ID
    
    
    retIdx = 1
   
    For i = LBound(uniqueGroupValues) To UBound(uniqueGroupValues)
        If uniqueGroupValues(i).ParentIndex = parentIdx Then
            ret(retIdx) = Table_GroupBy_Map_BuildGroupToRowMap(i, uniqueGroupValues)
            retIdx = retIdx + 1
        End If
    Next i
    
    
    Table_GroupBy_Map_BuildGroupToRowMap = ret

End Function
' Data_LoadFromRange: Loads column data as an NxM column array from a worksheet/range
Public Function Data_LoadFromRange(dataRng As Range, columns As Variant, Optional includeHeaderRow As Boolean = False, Optional preserveOrigArrayDimensions As Boolean = False) As Variant()
    
    ' 2015-06-12 (joe h): Added preserveOrigArrayDimensions and includeHeaderRow - which should be set true if data will be copied to spreadsheet range
    
    'Dim dataRng As Range: Set dataRng = Range("ALPT_BY_HR!A:AX")
    'Dim columns As Variant: columns = Array("ENODEB")
    
    Debug.Assert Not (dataRng Is Nothing)
    
    If Not IsArray(columns) Then columns = Array(columns)
    
    
    Dim i As Integer, j As Integer
    Dim lastRow As Long, numCols As Long, numEntries As Long, colNum As Integer
    Dim retData() As Variant
    Dim tmpData As Variant
    
    numCols = Array_Count(columns)
    
    lastRow = Range_LastRow(dataRng)
    numEntries = lastRow - IIf(includeHeaderRow, 0, 1) ' header row
    
    
    If preserveOrigArrayDimensions = False Then
        If numCols = 1 Then
            ReDim retData(1 To numEntries) As Variant
        Else
            ReDim retData(LBound(columns) To UBound(columns), 1 To numEntries) As Variant
        End If
    Else
        ReDim retData(1 To numEntries, LBound(columns) To UBound(columns)) As Variant
    End If
    
    For i = LBound(columns) To UBound(columns)
        
        If IsNumeric(columns(i)) Then
            colNum = columns(i)
        Else
            colNum = Range_GetColumnByName(dataRng, CStr(columns(i)))
            If colNum = 0 Then Err.Raise -1, , "Invalid column name: " & columns(i)
        End If
        
        ' Fast way to load data from worksheet
        tmpData = dataRng.Cells(IIf(includeHeaderRow, 1, 2), colNum).Resize(numEntries)
        
        If preserveOrigArrayDimensions = False Then
            For j = 1 To numEntries
                If numCols = 1 Then
                    ' convert 2D array to 1D array
                    retData(j) = tmpData(j, 1)
                Else
                    ' transpose
                    retData(i, j) = tmpData(j, 1)
                End If
            Next j
        Else
            For j = 1 To numEntries
                retData(j, i) = tmpData(j, 1)
            Next j
        End If
       
    Next i
    
    
    Data_LoadFromRange = retData

End Function
Public Function Data_LoadColumnDataToArray2_WIP() As Variant()


    Dim dataRng As Range: Set dataRng = Range("ALPT_DT!A:AX")
    
    Dim dataTest() As Variant, dti As Integer
    ReDim dataTest(1 To 4, 1 To 2) As Variant
    
    For dti = 1 To 4
        dataTest(dti, 1) = Choose(dti, "test_header1", 1, 2, 3)
        dataTest(dti, 2) = Choose(dti, "test_header2", 4 + 1, 4 + 2, 4 + 3)
    Next dti
    
   
    Dim dataSources As Variant
    dataSources = Array(Sheets("ALPT_DT"), "REGION", "MARKET", "SITE", dataTest, "test_header2", dataRng, "DAY", Range("ALPT_DT!G:I"), dataRng, "RRC_Conn_Fail_den")
    
    
    Dim i As Long, j As Long
    Dim dataSourcesExpanded As Variant
    
    dataSourcesExpanded = Array()
    
    'for i=lbound(datasourceexpanded) to ubound(
    

    
    
    Dim dsType As String, dataSourceIdx As Integer, dataSourceColCount As Long
    Dim colCount As Long, maxRows As Long
    Dim colNbr As Long, colRng As Range
    
    
    Dim ws As Worksheet, rng As Range
    Dim lastRow As Long, lastCol As Long
    Dim colNames() As Variant, colNums() As Long
    
    
    Dim tmpRng As Range, tmpData() As Variant, tmpCol As Long
    
    Dim data() As Variant
    Dim curCol As Long
    
    Dim passNo As Integer
    
    colCount = 0
    maxRows = 0
    
    
    For passNo = 1 To 2
    
        Debug.Print "Pass #: " & passNo
        Debug.Print "---------------------------------"
        
        ' Pass 1: Count the columns for each data source, and add. Determine maximum number of rows.
        ' Pass 2: Load the data
        
        If passNo = 2 Then
            ReDim data(1 To colCount, 1 To maxRows) As Variant
            
            curCol = 1
        End If
        
        
        dataSourceIdx = LBound(dataSources)
    
        Do While dataSourceIdx <= UBound(dataSources)
            dsType = TypeName(dataSources(dataSourceIdx))
            
            Debug.Print dsType
            
            If dsType = "Worksheet" Or dsType = "Range" Then
                If dsType = "Worksheet" Then
                    Set ws = dataSources(dataSourceIdx)
                    lastRow = Worksheet_GetLastRow(ws)
                    lastCol = Worksheet_GetLastColumn(ws)
                    
                    Set rng = ws.Cells(1, 1).Resize(lastRow, lastCol)
                ElseIf dsType = "Range" Then
                    Set rng = dataSources(dataSourceIdx)
                    
                    lastRow = Range_LastRow(rng)
                    lastCol = rng.columns.count
                End If
                
                dataSourceColCount = 0
                dataSourceIdx = dataSourceIdx + 1
                
                Do While dataSourceIdx <= UBound(dataSources)
                    If TypeName(dataSources(dataSourceIdx)) = "String" Then
                        Debug.Print "- " & dataSources(dataSourceIdx)
                        
                        If passNo = 2 Then
                            tmpCol = Range_GetColumnByName(rng, CStr(dataSources(dataSourceIdx)))
                            
                            If tmpCol > 0 Then
                                Set tmpRng = rng.Cells(1, tmpCol).Resize(lastRow)
                                tmpData = tmpRng
                                
                                For i = 1 To lastRow: data(curCol, i) = tmpData(i, 1): Next i
                                For i = lastRow + 1 To maxRows: data(curCol, i) = Null: Next i
                                
                            Else
                                Err.Raise -1, , "Cannot find column: " & dataSources(dataSourceIdx)
                            End If

                            Debug.Print "- curCol: " & curCol & "(+1)"
                            curCol = curCol + 1
                        End If
                        
                        
                        dataSourceColCount = dataSourceColCount + 1
                    Else
                        Exit Do
                    End If
                    
                    
                    dataSourceIdx = dataSourceIdx + 1
                Loop
                
                dataSourceIdx = dataSourceIdx - 1
                
                
                If passNo = 1 Then
                
                    ' Data source column count is 0 -> import all data in range or worksheet
                    If dataSourceColCount = 0 Then
                        dataSourceColCount = lastCol
                        Debug.Print "- All cols"
                    End If
                
                    colCount = colCount + dataSourceColCount
                    If lastRow > maxRows Then maxRows = lastRow
                    
                    Debug.Print "- Adding: " & dataSourceColCount
                
                ElseIf passNo = 2 Then
                
                    ' Data source column count is 0 -> import all data in range or worksheet
                    If dataSourceColCount = 0 Then
                        tmpData = rng
                        
                        For j = 1 To lastCol
                            For i = 1 To lastRow: data(curCol, i) = tmpData(i, j): Next i
                            For i = lastRow + 1 To maxRows: data(curCol, i) = Null: Next i
                            
                            Debug.Print "- curCol: " & curCol & "(+1)"
                            curCol = curCol + 1
                            
                        Next j
                    End If
                
                End If
                
                
            ElseIf IsArray(dataSources(dataSourceIdx)) Then
            
                If Array_NumDimensions(dataSources(dataSourceIdx)) = 2 Then
                    Dim arrayData() As Variant
                    
                    arrayData = dataSources(dataSourceIdx)
                
                    lastRow = UBound(arrayData, 1) - LBound(arrayData, 1) + 1
                    lastCol = UBound(arrayData, 2) - LBound(arrayData, 2) + 1
                    
                    
                
                    dataSourceColCount = 0
                    dataSourceIdx = dataSourceIdx + 1
                    
                    Do While dataSourceIdx <= UBound(dataSources)
                        If TypeName(dataSources(dataSourceIdx)) = "String" Then
                            Debug.Print "- " & dataSources(dataSourceIdx)
                        
                            If passNo = 2 Then
                                Dim colIdx As Long, colName As String
                                
                                colIdx = -1
                                colName = CStr(dataSources(dataSourceIdx))
                                
                                For i = LBound(arrayData, 2) To UBound(arrayData, 2)
                                    If arrayData(1, i) = colName Then
                                        colIdx = i
                                        Exit For
                                    End If
                                Next i
                                
                                
                                
                                If colIdx > 0 Then
                                    For i = 1 To lastRow: data(curCol, i) = arrayData(i, colIdx): Next i
                                    For i = lastRow + 1 To maxRows: data(curCol, i) = Null: Next i
                                    
                                Else
                                    Err.Raise -1, , "Cannot find column: " & colName
                                End If
    
                                Debug.Print "- curCol: " & curCol & "(+1)"
                                curCol = curCol + 1
                            End If
                            
                            
                            dataSourceColCount = dataSourceColCount + 1
                        Else
                            Exit Do
                        End If
                        
                        
                        dataSourceIdx = dataSourceIdx + 1
                    Loop
                    dataSourceIdx = dataSourceIdx - 1
                    
                    If passNo = 1 Then
                        If dataSourceColCount = 0 Then
                            dataSourceColCount = lastCol
                            Debug.Print "- All cols"
                        End If
                        
                        colCount = colCount + dataSourceColCount
                        If lastRow > maxRows Then maxRows = lastRow
                        
                        
                        Debug.Print "- Adding: " & dataSourceColCount
                    ElseIf passNo = 2 Then
                    
                        If dataSourceColCount = 0 Then
                            For j = LBound(arrayData, 2) To UBound(arrayData, 2)
                                For i = 1 To lastRow: data(curCol, i) = arrayData(i, j): Next i
                                For i = lastRow + 1 To maxRows: data(curCol, i) = Null: Next i
                                
                                curCol = curCol + 1
                            Next j
                        End If
                    End If
                    
                    
                End If
                
            End If
        
            dataSourceIdx = dataSourceIdx + 1
        Loop
    
    Next passNo
    
    
    Data_LoadColumnDataToArray2_WIP = data
    
    Debug.Print
    
    Exit Function
    
    
    
    ReDim data(1 To colCount, 1 To maxRows) As Variant
    
    dataSourceIdx = LBound(dataSources)
    curCol = 1
    
    Do While dataSourceIdx <= UBound(dataSources)
        dsType = TypeName(dataSources(dataSourceIdx))
        
        
        If dsType = "Worksheet" Or dsType = "Range" Then
            If dsType = "Worksheet" Then
                Set ws = dataSources(dataSourceIdx)
                lastRow = Worksheet_GetLastRow(ws)
                lastCol = Worksheet_GetLastColumn(ws)
                
                Set rng = ws.Cells(1, 1).Resize(lastRow, lastCol)
            ElseIf dsType = "Range" Then
                Set rng = dataSources(dataSourceIdx)
                
                lastRow = Range_LastRow(rng)
                lastCol = rng.columns.count
            End If
            
            dataSourceIdx = dataSourceIdx + 1
            dataSourceColCount = 0
            
            Do While dataSourceIdx <= UBound(dataSources)
                If TypeName(dataSources(dataSourceIdx)) = "String" Then
                    tmpCol = Range_GetColumnByName(rng, CStr(dataSources(dataSourceIdx)))
                    
                    If tmpCol > 0 Then
                        Set tmpRng = rng.Cells(1, tmpCol).Resize(lastRow)
                        tmpData = tmpRng
                        
                        For i = 1 To lastRow: data(curCol, i) = tmpData(i, 1): Next i
                        For i = lastRow + 1 To maxRows: data(curCol, i) = Null: Next i
                        
                    Else
                        Err.Raise -1, , "Cannot find column: " & dataSources(dataSourceIdx)
                    End If
                
                    dataSourceColCount = dataSourceColCount + 1
                    curCol = curCol + 1
                Else
                    Exit Do
                End If
                
                dataSourceIdx = dataSourceIdx + 1
            Loop
            
            dataSourceIdx = dataSourceIdx - 1
            
            
            
            ' Data source column count is 0 -> import all data in range or worksheet
            If dataSourceColCount = 0 Then
                tmpData = rng
                
                For j = 1 To lastCol
                    For i = 1 To lastRow: data(curCol, i) = tmpData(i, j): Next i
                    For i = lastRow + 1 To maxRows: data(curCol, i) = Null: Next i
                    
                    curCol = curCol + 1
                Next j
            End If
    
        ElseIf IsArray(dataSources(dataSourceIdx)) Then
        
            If Array_NumDimensions(dataSources(dataSourceIdx)) = 2 Then
                lastRow = UBound(dataSources(dataSourceIdx), 1) - LBound(dataSources(dataSourceIdx), 1) + 1
                lastCol = UBound(dataSources(dataSourceIdx), 2) - LBound(dataSources(dataSourceIdx), 2) + 1
                
                
                
                
                'tmpData = dataSources(dataSourceIdx)
                
                'For j = 1 To lastCol
                '    For i = 1 To dataSourceRows: data(curCol, i) = tmpData(i, j): Next i
                '    For i = dataSourceRows + 1 To maxRows: data(curCol, i) = Null: Next i
                '
                '    curCol = curCol + 1
                'Next j
                
                
            End If
            
        End If
    
        dataSourceIdx = dataSourceIdx + 1
    Loop
    
    
    Debug.Print

End Function

Public Function Data_UniqueSets_WIP() '(data() As Variant, Optional addCountColumn As Boolean = False) As Variant()

    ' work in progress

    Dim data() As Variant: data = Data_LoadColumnDataToArray2_WIP
    Dim groupByColumns As Variant: groupByColumns = Array()
    Dim addCountColumn As Boolean: addCountColumn = True
    
    If Array_NumDimensions(data) <> 2 Then Err.Raise -1, , "Error: Parameter columnData must be an 2D array"
    
    Dim i As Long, j As Long, k As Long
    
    Dim dataCols() As Long, colIdx As Long
    
    For i = LBound(groupByColumns) To UBound(groupByColumns)
    Next i
    
    
    
    
    
    Dim lbColumns As Integer, ubColumns As Integer
    Dim lbDataRows As Long, ubDataRows As Long
    Dim numEntries As Long
    
    Dim test As Variant
    
    lbColumns = LBound(data, 1)
    ubColumns = UBound(data, 1)
    lbDataRows = LBound(data, 2)
    ubDataRows = UBound(data, 2)
    
    
    
    Dim uniqueSets() As Variant
    Dim uniqueSets_N() As Long
    Dim uniqueSetCount As Long
    Dim uniqueSetIdx As Long
    
    
    Dim lbUniqueSetCols As Integer, ubUniqueSetCols As Integer
    
    lbUniqueSetCols = lbColumns
    ubUniqueSetCols = ubColumns + IIf(addCountColumn = True, 1, 0)
    

    uniqueSetCount = 0

    For i = lbDataRows To ubDataRows
        uniqueSetIdx = 0
        
        For j = 1 To uniqueSetCount
            uniqueSetIdx = j
            
            For k = lbColumns To ubColumns
                If Not data(k, i) = uniqueSets(k, j) Then
                    uniqueSetIdx = 0
                    Exit For
                End If
            Next k
            
            ' Found unique set
            If uniqueSetIdx > 0 Then Exit For
        Next j
        
        If uniqueSetIdx = 0 Then
            uniqueSetCount = uniqueSetCount + 1
            uniqueSetIdx = uniqueSetCount
            
            ReDim Preserve uniqueSets(lbUniqueSetCols To ubUniqueSetCols, 1 To uniqueSetCount) As Variant
            ReDim Preserve uniqueSets_N(uniqueSetCount) As Long
            
            For k = lbColumns To ubColumns
                uniqueSets(k, uniqueSetCount) = data(k, i)
            Next k
            
            
            
            uniqueSets_N(uniqueSetIdx) = 0
            
        End If
        
        uniqueSets_N(uniqueSetIdx) = uniqueSets_N(uniqueSetIdx) + 1
        
        ' Last Column contains the number of entries per unique set (aka count)
        If addCountColumn Then uniqueSets(ubColumns + 1, uniqueSetIdx) = uniqueSets(ubColumns + 1, uniqueSetIdx) + 1

    Next i
    
    Data_UniqueSets_WIP = uniqueSets

End Function


Public Function Data_CalcAggrKpis_WIP()

    ' This is a work in progress

    Dim dataRng As Range
    Dim groupByFields() As Variant
    Dim valueFields() As Variant
    Dim valueFuncs() As Variant
    Dim aggrFieldNames() As Variant
    Dim optFirstRowColHeaders As Boolean

    
    
    optFirstRowColHeaders = False
    
    Set dataRng = Range("ALPT_BY_HR!A:AX")
    groupByFields = Array("DAY", "ENODEB", "EUTRANCELL", "CARRIER") ' Array("DAY", "ENODEB", "EUTRANCELL", "CARRIER")
    
    valueFields = Array("RRC_Conn_Fail_den") 'RRC_Conn_Fail_den
    valueFuncs = Array("SUM")
    'aggrFieldNames = Array("SUM(RRC_Conn_Fail_den)")
    
    
    Debug.Assert UBound(valueFields) > -1
    Debug.Assert LBound(valueFields) = LBound(valueFuncs)
    Debug.Assert UBound(valueFields) = UBound(valueFuncs)
    
    
    Dim i As Integer, j As Integer, k As Integer
    Dim lastRow As Long, numEntries As Long
    
    lastRow = Range_LastRow(dataRng)
    
    numEntries = lastRow - 1
    
    Dim tmpData As Variant, colNum As Integer
    Dim groupByData() As Variant, valueFieldData() As Variant
    Dim groupByFieldCount As Integer, aggrFieldCount As Integer
    'Dim intermediateValues() As AggregateFuncIntermediateValues
    
    groupByFieldCount = Array_Count(groupByFields)
    aggrFieldCount = Array_Count(valueFields)
    
    
    Debug.Assert groupByFieldCount > 0
    
    
    ' --------------------------------------------------------------------------------------
    ' Begin - Load Data from Group By fields
    ' --------------------------------------------------------------------------------------
    'If groupByFieldCount > 0 Then
        ReDim groupByData(LBound(groupByFields) To UBound(groupByFields), 1 To numEntries) As Variant
        
        For i = LBound(groupByFields) To UBound(groupByFields)
            
            If IsNumeric(groupByFields(i)) Then
                colNum = groupByFields(i)
            Else
                colNum = Range_GetColumnByName(dataRng, CStr(groupByFields(i)))
                If colNum = 0 Then Err.Raise -1, , "Invalid column name: " & groupByFields(i)
            End If
            
            ' Fast way to load data from worksheet
            tmpData = dataRng.Cells(2, colNum).Resize(numEntries)
            For j = 1 To numEntries: groupByData(i, j) = tmpData(j, 1): Next j ' convert 2D array to 1D array
           
        Next i
    'End If
    ' --------------------------------------------------------------------------------------
    ' End - Load Data from Group By fields
    ' --------------------------------------------------------------------------------------
    
    
    ' --------------------------------------------------------------------------------------
    ' Begin - Load Data from value fields
    ' --------------------------------------------------------------------------------------
    ReDim valueFieldData(LBound(valueFields) To UBound(valueFields), 1 To numEntries) As Variant
    

    For i = LBound(valueFields) To UBound(valueFields)
        If IsNumeric(valueFields(i)) Then
            colNum = valueFields(i)
        Else
            colNum = Range_GetColumnByName(dataRng, CStr(valueFields(i)))
            If colNum = 0 Then Err.Raise -1, , "Invalid column name: " & valueFields(i)
        End If
        
        ' Fast way to load data from worksheet
        tmpData = dataRng.Cells(2, colNum).Resize(numEntries)
        For j = 1 To numEntries: valueFieldData(i, j) = tmpData(j, 1): Next j ' convert 2D  array to 1D array
        
    Next i
    ' --------------------------------------------------------------------------------------
    ' Begin - Load Data from value fields
    ' --------------------------------------------------------------------------------------




    Dim groupByUniqueSets() As Variant
    Dim groupByUniqueSets_N() As Long
    Dim groupByUniqueSetCount As Long
    Dim groupByUniqueSetIdx As Long
    
    Dim intermediateValues_N() As Long
    Dim intermediateValues_Value() As Variant
    

    groupByUniqueSetCount = 0

    For i = 1 To numEntries
        groupByUniqueSetIdx = 0
        
        For j = 1 To groupByUniqueSetCount
            groupByUniqueSetIdx = j
            
            For k = LBound(groupByFields) To UBound(groupByFields)
                If Not groupByData(k, i) = groupByUniqueSets(k, j) Then
                    groupByUniqueSetIdx = 0
                    Exit For
                End If
            Next k
            
            ' Found unique set
            If groupByUniqueSetIdx > 0 Then Exit For
        Next j
        
        If groupByUniqueSetIdx = 0 Then
            groupByUniqueSetCount = groupByUniqueSetCount + 1
            groupByUniqueSetIdx = groupByUniqueSetCount
            
            ReDim Preserve groupByUniqueSets(LBound(groupByFields) To UBound(groupByFields), 1 To groupByUniqueSetCount) As Variant
            
            For k = LBound(groupByFields) To UBound(groupByFields)
                groupByUniqueSets(k, groupByUniqueSetCount) = groupByData(k, i)
            Next k
            
            
            ReDim Preserve intermediateValues_N(1 To groupByUniqueSetCount) As Long
            ReDim Preserve intermediateValues_Value(LBound(valueFields) To UBound(valueFields), 1 To groupByUniqueSetCount) As Variant
            
            intermediateValues_N(groupByUniqueSetCount) = 0
            
            For j = LBound(valueFields) To UBound(valueFields): intermediateValues_Value(j, groupByUniqueSetCount) = vbNull: Next j
            
        End If
        
        intermediateValues_N(groupByUniqueSetIdx) = intermediateValues_N(groupByUniqueSetIdx) + 1


        For j = LBound(valueFields) To UBound(valueFields)
            Select Case valueFuncs(j)
                Case "SUM":
                    If intermediateValues_Value(j, groupByUniqueSetIdx) = vbNull Then
                        intermediateValues_Value(j, groupByUniqueSetIdx) = valueFieldData(j, i)
                    Else
                        intermediateValues_Value(j, groupByUniqueSetIdx) = intermediateValues_Value(j, groupByUniqueSetIdx) + valueFieldData(j, i)
                    End If
                Case "MAX":
                    If intermediateValues_Value(j, groupByUniqueSetIdx) = vbNull Or valueFieldData(j, i) > intermediateValues_Value(j, groupByUniqueSetIdx) Then
                        intermediateValues_Value(j, groupByUniqueSetIdx) = valueFieldData(j, i)
                    End If
                Case "MIN":
                    If intermediateValues_Value(j, groupByUniqueSetIdx) = vbNull Or valueFieldData(j, i) < intermediateValues_Value(j, groupByUniqueSetIdx) Then
                        intermediateValues_Value(j, groupByUniqueSetIdx) = valueFieldData(j, i)
                    End If
            End Select
        Next j
        
    Next i
    
    
    Dim retArray() As Variant
    Dim retFieldCount As Integer
    Dim retHeaderOffset As Integer
    
    
    retHeaderOffset = 0
    retFieldCount = groupByFieldCount + aggrFieldCount
    
    ReDim retArray(1 To groupByUniqueSetCount + retHeaderOffset, 1 To retFieldCount) As Variant
    
    If optFirstRowColHeaders Then
        For j = LBound(groupByFields) To UBound(groupByFields)
            retArray(1, j - LBound(groupByFields) + 1) = groupByFields(j)
        Next j
        For j = LBound(valueFields) To UBound(valueFields)
            retArray(1, j - LBound(valueFields) + groupByFieldCount + 1) = valueFuncs(j) & "(" & valueFields(j) & ")"
        Next j
        
        retHeaderOffset = 1
    End If
    
    For i = 1 To groupByUniqueSetCount
        For j = LBound(groupByFields) To UBound(groupByFields)
            retArray(i + retHeaderOffset, j - LBound(groupByFields) + 1) = groupByUniqueSets(j, i)
        Next j
        For j = LBound(valueFields) To UBound(valueFields)
            retArray(i + retHeaderOffset, j - LBound(valueFields) + groupByFieldCount + 1) = intermediateValues_Value(j, i)
        Next j
    Next i
    
    Debug.Print

End Function
Public Sub Debug_Init()
    
    
    ReDim Debug_Lines(1 To 100) As String
   
    Debug_CurLine = 0
    Debug_Capacity = 100
    Debug_SilentFlag = False

End Sub
Public Sub Debug_Print(str As String)

    If Not Debug_SilentFlag Then
        If Debug_CurLine = Debug_Capacity Then
            Debug_Capacity = Debug_Capacity * 2
            ReDim Preserve Debug_Lines(1 To Debug_Capacity)
        End If
        
        Debug_CurLine = Debug_CurLine + 1
        
        Debug_Lines(Debug_CurLine) = str
    End If
   
    
End Sub
Public Sub Debug_Silent(flg As Boolean)
    
    Debug_SilentFlag = flg

End Sub
Public Sub Debug_ToExcel(Optional sheet As String = "DEBUG")

    Dim prevAppScreenUpdating As Boolean
    Dim ws As Worksheet
    Dim i As Long
    
    
    prevAppScreenUpdating = Application.ScreenUpdating
    Application.ScreenUpdating = False

    
    If Worksheet_SheetExists(sheet) Then
        Set ws = ThisWorkbook.Worksheets(sheet)
        ws.Cells.Clear
    Else
        Set ws = ThisWorkbook.Sheets.Add()
        ws.Name = sheet
    End If
    
    
    For i = 1 To Debug_CurLine
        ws.Cells(i, 1).Value = Debug_Lines(i)
    Next i
    
    
    Application.ScreenUpdating = prevAppScreenUpdating

End Sub


Public Sub Output_Init()

    
    ReDim Output_Lines(1 To 100) As String
   
    Output_CurLine = 0
    Output_Capacity = 100
    Output_SilentFlag = False
    
    Output_CM_NumCols = 0
    
    Output_IndentLevel = 0

End Sub
Public Sub Output_Print(str As String)

    Dim i As Integer

    If Not Output_SilentFlag Then
    
        If Output_CurLine = Output_Capacity Then
            Output_Capacity = Output_Capacity * 2
            ReDim Preserve Output_Lines(1 To Output_Capacity)
        End If
        
        Output_CurLine = Output_CurLine + 1
        
        For i = 1 To Output_IndentLevel
            Output_Lines(Output_CurLine) = Output_Lines(Output_CurLine) & "  "
        Next i
        
        Output_Lines(Output_CurLine) = Output_Lines(Output_CurLine) & str
    End If
   
    
End Sub
Public Sub Output_Silent(flg As Boolean)
    
    Output_SilentFlag = flg

End Sub
Public Sub Output_Indent_Increase()

    Output_IndentLevel = Output_IndentLevel + 1

End Sub
Public Sub Output_Indent_Decrease()

    Output_IndentLevel = IIf(Output_IndentLevel > 0, Output_IndentLevel - 1, 0)

End Sub
Public Sub Output_ToExcel(Optional sheet As String = "OUTPUT")

    Dim prevAppScreenUpdating As Boolean
    Dim ws As Worksheet
    Dim i As Long
    
    
    
    If Worksheet_SheetExists(sheet) Then
        Set ws = ThisWorkbook.Worksheets(sheet)
        ws.Cells.Clear
    Else
        Set ws = ThisWorkbook.Sheets.Add()
        ws.Name = sheet
    End If
    
    prevAppScreenUpdating = Application.ScreenUpdating
    Application.ScreenUpdating = False
    
    For i = 1 To Output_CurLine
        ws.Cells(i, 1).Value = Output_Lines(i)
    Next i
    
    
    Application.ScreenUpdating = prevAppScreenUpdating

End Sub
Public Sub Output_CM_Start(Optional numColumns As Integer = 2)
    
    Debug.Assert numColumns > 1
    
    Output_CM_NumCols = numColumns
    
    Output_CM_CapacityPerCol = 100
    
    ReDim Output_CM_Lines(1 To Output_CM_CapacityPerCol * Output_CM_NumCols) As String
    ReDim Output_CM_CurLines(1 To Output_CM_NumCols) As Long
   
    Dim i As Integer
    
    For i = 1 To Output_CM_NumCols
        Output_CM_CurLines(i) = 0
    Next i
    
    Output_CM_HBar_Count = 0
    
End Sub
Public Sub Output_CM_Print(colNum As Integer, str As String)

    Debug.Assert colNum >= 1
    Debug.Assert colNum <= Output_CM_NumCols

    If Not Output_SilentFlag Then
        If Output_CM_CurLines(colNum) = Output_CM_CapacityPerCol Then
            'Output_Capacity = Output_Capacity + 100
            'ReDim Preserve Output_Lines(1 To Output_Capacity)
            Output_Print "Reached output col capacity"
            Exit Sub
        End If
        
        Output_CM_CurLines(colNum) = Output_CM_CurLines(colNum) + 1
        
        Output_CM_Lines((colNum - 1) * Output_CM_CapacityPerCol + Output_CM_CurLines(colNum)) = str
    End If
    
End Sub
Public Sub Output_CM_HBar()

    Debug.Assert Output_CM_NumCols > 0
    
    Output_CM_HBar_Count = Output_CM_HBar_Count + 1

    ReDim Preserve Output_CM_HBar_Lines(1 To Output_CM_HBar_Count) As Integer
    
    Dim col As Integer, maxLines As Integer
    
    maxLines = 0
    
    For col = 1 To Output_CM_NumCols
        If Output_CM_CurLines(col) > maxLines Then
            maxLines = Output_CM_CurLines(col)
        End If
    Next col
    
    Output_CM_HBar_Lines(Output_CM_HBar_Count) = maxLines

End Sub
Public Sub Output_CM_EndFlush(Optional minColWidth As Integer = 0, Optional colSep As String = " ")

    Debug.Assert Output_CM_NumCols > 0

    Dim col As Integer, lineNo As Integer
    Dim i As Integer, j As Integer
    
    Dim maxLines As Integer
    Dim maxLineWidthPerCol() As Integer
    Dim colLineStr As String
    Dim outStr As String
    Dim paddingSize As Integer
    
    Dim hbar_cur As Integer
    
    ReDim maxLineWidthPerCol(1 To Output_CM_NumCols) As Integer
    
    maxLines = 0
    
    For col = 1 To Output_CM_NumCols
        If Output_CM_CurLines(col) > maxLines Then
            maxLines = Output_CM_CurLines(col)
        End If
        
        maxLineWidthPerCol(col) = 0
        
        For lineNo = 1 To Output_CM_CurLines(col)
            colLineStr = Output_CM_Lines((col - 1) * Output_CM_CapacityPerCol + lineNo)
            
            If Len(colLineStr) > maxLineWidthPerCol(col) Then
                maxLineWidthPerCol(col) = Len(colLineStr)
            End If
        Next lineNo
    Next col
    
    hbar_cur = 1
    
    
    
    For lineNo = 1 To maxLines
        
        If Output_CM_HBar_Count > 0 And hbar_cur <= Output_CM_HBar_Count Then
            If Output_CM_HBar_Lines(hbar_cur) < lineNo Then
                paddingSize = 0
            
                outStr = ""
                
                For col = 1 To Output_CM_NumCols
                    paddingSize = IIf(maxLineWidthPerCol(col) > minColWidth, maxLineWidthPerCol(col), minColWidth)
                    
                    paddingSize = paddingSize + IIf(col < Output_CM_NumCols, Len(colSep) - IIf(col = 1, Len(LTrim(colSep)), Len(Trim(colSep))), 0)
                    
                    For j = 1 To paddingSize
                        outStr = outStr & "-"
                    Next j
                    
                    outStr = outStr + IIf(col < Output_CM_NumCols, Trim(colSep), "")
                    
                Next col
                
                'For col = 1 To Output_CM_NumCols
                '    paddingSize = paddingSize + IIf(maxLineWidthPerCol(col) > minColWidth, maxLineWidthPerCol(col), minColWidth)
                '    paddingSize = paddingSize + IIf(col < Output_CM_NumCols, Len(colSep), 0)
                'Next col
                
                'outStr = ""
                
                'For i = 1 To paddingSize
                '    outStr = outStr & "-"
                'Next i
                
                
                Output_Print outStr
                
                hbar_cur = hbar_cur + 1
                
            End If
        End If
        
        outStr = ""
        
        For col = 1 To Output_CM_NumCols
            If Output_CM_CurLines(col) <= maxLines Then
                colLineStr = Output_CM_Lines((col - 1) * Output_CM_CapacityPerCol + lineNo)
                
                outStr = outStr & colLineStr
                
                paddingSize = IIf(maxLineWidthPerCol(col) > minColWidth, maxLineWidthPerCol(col), minColWidth) - Len(colLineStr)
                
                For i = 1 To paddingSize
                    outStr = outStr & " "
                Next i
                
                outStr = outStr & IIf(col < Output_CM_NumCols, colSep, "")
            End If
        Next col
        
            
        Output_Print outStr
    Next lineNo
    
    Output_CM_NumCols = 0

End Sub
