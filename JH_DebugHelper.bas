Attribute VB_Name = "JH_DebugHelper"
Option Explicit
Option Private Module


' JH_DebugHelper
' ------------------------------------------------------------------------------------------------------------------
' Module to help buffer output for debug and/or output purposes
'
' Joseph Huntley (joseph.huntley@vzw.com)
' ------------------------------------------------------------------------------------------------------------------
' Changelog:
'
' 2017-10-02 (joe h)    - Added OutputBuffer_CM_PrintRow, and corresponding Debug_XX function
'                       - Fixed OutputBuffer_CM_EndFlush not printing out HBars added after the last row
' 2017-09-29 (joe h)    - Added Debug_VariableListString
' 2017-09-20 (joe h)    - Created OutputBuffer type to used in encapsuling function (e.g. Debug_)


Private Type OutputBuffer
    InitFlag As Boolean
    CurLine As Long
    Lines() As String
    Capacity As Long
    SilentFlag As Boolean
    IndentLevel As Integer
    
    ' Column Matrix variables
    CM_CurLines() As Long
    CM_Lines() As String
    CM_CapacityPerCol As Long
    CM_NumCols As Integer
    CM_HBar_Lines() As Integer
    CM_HBar_Count As Integer
End Type

Private obDebug As OutputBuffer ' Output buffer for Debug_ functions

Private Sub OutputBuffer_Init(ob As OutputBuffer, Optional reset As Boolean = False)

    If ob.InitFlag And Not reset Then Exit Sub

    ReDim ob.Lines(1 To 100) As String
   
    ob.CurLine = 0
    ob.Capacity = 100
    ob.SilentFlag = False
    
    ob.CM_NumCols = 0
    
    ob.IndentLevel = 0
    
    ob.InitFlag = True

End Sub
Private Function OutputBuffer_Print(ob As OutputBuffer, ByVal outputToPrint As Variant) As String


    OutputBuffer_Init ob

    Dim i As Integer

    If ob.SilentFlag Then Exit Function
    
    ' If we're at capacity, increase the lines
    If ob.CurLine = ob.Capacity Then
        ob.Capacity = ob.Capacity * 2
        ReDim Preserve ob.Lines(1 To ob.Capacity)
    End If
    
    ob.CurLine = ob.CurLine + 1
    
    For i = 1 To ob.IndentLevel
        ob.Lines(ob.CurLine) = ob.Lines(ob.CurLine) & "  "
    Next i
        
    Dim output As Variant, addTab As Boolean
    
    If Not IsArray(outputToPrint) Then outputToPrint = Array(outputToPrint)
    
    addTab = False
    
    For Each output In outputToPrint
        ob.Lines(ob.CurLine) = ob.Lines(ob.CurLine) & IIf(addTab, vbTab, "") & output
        addTab = True
    Next output
    
    OutputBuffer_Print = ob.Lines(ob.CurLine)
    
End Function
Private Sub OutputBuffer_Silent(ob As OutputBuffer, flg As Boolean)
    ob.SilentFlag = flg
End Sub
Private Sub OutputBuffer_Indent_Increase(ob As OutputBuffer)
    ob.IndentLevel = ob.IndentLevel + 1
End Sub
Private Sub OutputBuffer_Indent_Decrease(ob As OutputBuffer)
    ob.IndentLevel = IIf(ob.IndentLevel > 0, ob.IndentLevel - 1, 0)
End Sub
Private Sub OutputBuffer_ToExcel(ob As OutputBuffer, Optional sheet As String = "OUTPUT")

    Dim prevAppScreenUpdating As Boolean
    Dim ws As Worksheet
    Dim i As Long
    
    
    If ob.CurLine = 0 Then Exit Sub
    
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheet)
    
    If Err.Number <> 0 Then
        ws.Cells.Clear
    Else
        Set ws = ThisWorkbook.Sheets.Add()
        ws.Name = sheet
    End If
    On Error GoTo 0
    
    
    prevAppScreenUpdating = Application.ScreenUpdating
    Application.ScreenUpdating = False
    
    For i = 1 To ob.CurLine
        ws.Cells(i, 1).Value = ob.Lines(i)
    Next i
    
    
    Application.ScreenUpdating = prevAppScreenUpdating

End Sub
Private Sub OutputBuffer_ToFile(ob As OutputBuffer, filePath As String)

    
    If ob.CurLine = 0 Then Exit Sub

    Dim fileHandle As Long, i As Long
    
    fileHandle = FreeFile
    
    Open filePath For Output As #fileHandle
    
        For i = 1 To ob.CurLine
            Print #fileHandle, ob.Lines(i)
        Next i
    Close #fileHandle
    

End Sub
Private Sub OutputBuffer_CM_Start(ob As OutputBuffer, Optional numColumns As Integer = 2)
    
    Debug.Assert numColumns > 1
    
    ob.CM_NumCols = numColumns
    
    ob.CM_CapacityPerCol = 100
    
    ReDim ob.CM_Lines(1 To ob.CM_CapacityPerCol * ob.CM_NumCols) As String
    ReDim ob.CM_CurLines(1 To ob.CM_NumCols) As Long
   
    Dim i As Integer
    
    For i = 1 To ob.CM_NumCols
        ob.CM_CurLines(i) = 0
    Next i
    
    ob.CM_HBar_Count = 0
    
End Sub
Private Sub OutputBuffer_CM_Print(ob As OutputBuffer, colNum As Integer, str As String)

    Debug.Assert colNum >= 1
    Debug.Assert colNum <= ob.CM_NumCols

    If Not ob.SilentFlag Then
        If ob.CM_CurLines(colNum) = ob.CM_CapacityPerCol Then
            'ob.Capacity = ob.Capacity + 100
            'ReDim Preserve ob.Lines(1 To ob.Capacity)
            OutputBuffer_Print ob, "Reached output col capacity"
            Exit Sub
        End If
        
        ob.CM_CurLines(colNum) = ob.CM_CurLines(colNum) + 1
        
        ob.CM_Lines((colNum - 1) * ob.CM_CapacityPerCol + ob.CM_CurLines(colNum)) = str
    End If

    
End Sub
Private Sub OutputBuffer_CM_PrintRow(ob As OutputBuffer, ByVal columnValues As Variant)
    
    Dim colCount As Long
    
    
    If Not IsArray(columnValues) Then columnValues = Array(columnValues)
    
    colCount = UBound(columnValues) - LBound(columnValues) = 1
    
    Debug.Assert colCount <= ob.CM_NumCols ' Will fail if array size is greater than allocated columns
    
    
    Dim i As Long
    
    For i = LBound(columnValues) To UBound(columnValues)
        OutputBuffer_CM_Print ob, i - LBound(columnValues) + 1, CStr(columnValues(i))
    Next i
    

End Sub
Private Sub OutputBuffer_CM_HBar(ob As OutputBuffer)

    Debug.Assert ob.CM_NumCols > 0
    
    ob.CM_HBar_Count = ob.CM_HBar_Count + 1

    ReDim Preserve ob.CM_HBar_Lines(1 To ob.CM_HBar_Count) As Integer
    
    Dim col As Integer, maxLines As Integer
    
    maxLines = 0
    
    For col = 1 To ob.CM_NumCols
        If ob.CM_CurLines(col) > maxLines Then
            maxLines = ob.CM_CurLines(col)
        End If
    Next col
    
    ob.CM_HBar_Lines(ob.CM_HBar_Count) = maxLines

End Sub
Private Sub OutputBuffer_CM_EndFlush(ob As OutputBuffer, Optional minColWidth As Integer = 0, Optional colSep As String = "|", Optional linePrefix As String = "|", Optional lineSuffix As String = "|")

    Debug.Assert ob.CM_NumCols > 0
    
    
    OutputBuffer_Init ob

    Dim col As Integer, lineNo As Integer
    Dim i As Integer, j As Integer
    
    Dim maxLines As Integer
    Dim maxLineWidthPerCol() As Integer
    Dim colLineStr As String
    Dim outStr As String
    Dim paddingSize As Integer
    
    Dim hbar_cur As Integer
    
    ReDim maxLineWidthPerCol(1 To ob.CM_NumCols) As Integer
    
    maxLines = 0
    
    For col = 1 To ob.CM_NumCols
        If ob.CM_CurLines(col) > maxLines Then
            maxLines = ob.CM_CurLines(col)
        End If
        
        maxLineWidthPerCol(col) = 0
        
        For lineNo = 1 To ob.CM_CurLines(col)
            colLineStr = ob.CM_Lines((col - 1) * ob.CM_CapacityPerCol + lineNo)
            
            If Len(colLineStr) > maxLineWidthPerCol(col) Then
                maxLineWidthPerCol(col) = Len(colLineStr)
            End If
        Next lineNo
    Next col
    
    hbar_cur = 1
    
    ' Generator HBar Line String
    Dim hbar_line_str As String
    
    hbar_line_str = linePrefix
    For col = 1 To ob.CM_NumCols
        paddingSize = IIf(maxLineWidthPerCol(col) > minColWidth, maxLineWidthPerCol(col), minColWidth)
        paddingSize = paddingSize + IIf(col < ob.CM_NumCols, Len(colSep) - IIf(col = 1, Len(LTrim(colSep)), Len(Trim(colSep))), 0)
        
        hbar_line_str = hbar_line_str & String(paddingSize, "-")
        'hbar_line_str = hbar_line_str & IIf(col < ob.CM_NumCols, String(Len(Trim(colSep)), "-"), "") ' this line inserts a dash between columns
        hbar_line_str = hbar_line_str & IIf(col < ob.CM_NumCols, Trim(colSep), "")                     ' this line inserts the column separator between columns
    Next col
    hbar_line_str = hbar_line_str & lineSuffix
    
    
    For lineNo = 1 To maxLines
        
        If ob.CM_HBar_Count > 0 And hbar_cur <= ob.CM_HBar_Count Then
            If ob.CM_HBar_Lines(hbar_cur) < lineNo Then
                OutputBuffer_Print ob, hbar_line_str
                hbar_cur = hbar_cur + 1
            End If
        End If
        
        outStr = linePrefix
        
        For col = 1 To ob.CM_NumCols
            If ob.CM_CurLines(col) <= maxLines Then
                colLineStr = ob.CM_Lines((col - 1) * ob.CM_CapacityPerCol + lineNo)
                
                outStr = outStr & colLineStr
                
                paddingSize = IIf(maxLineWidthPerCol(col) > minColWidth, maxLineWidthPerCol(col), minColWidth) - Len(colLineStr)
                
                
                For i = 1 To paddingSize
                    outStr = outStr & " "
                Next i
                
                outStr = outStr & IIf(col < ob.CM_NumCols, colSep, "")
            End If
        Next col
        
        
        outStr = outStr & lineSuffix
       
        OutputBuffer_Print ob, outStr
    Next lineNo
    
    ' Print out any HBars at end
    Do While hbar_cur <= ob.CM_HBar_Count
        OutputBuffer_Print ob, hbar_line_str
        hbar_cur = hbar_cur + 1
    Loop
    
    ob.CM_NumCols = 0

End Sub
Public Sub Debug_Init(Optional reset As Boolean = False)
    OutputBuffer_Init obDebug, reset
End Sub
Public Sub Debug_Print(ParamArray outputArr() As Variant)
    Dim printStr As String
    printStr = OutputBuffer_Print(obDebug, outputArr)
    
    ' Print to immediate window
    If Len(printStr) > 0 Then Debug.Print printStr
End Sub
Public Sub Debug_Silent(flg As Boolean)
    OutputBuffer_Silent obDebug, flg
End Sub
Public Sub Debug_Indent_Increase()
    OutputBuffer_Indent_Increase obDebug
End Sub
Public Sub Debug_Indent_Decrease()
    OutputBuffer_Indent_Decrease obDebug
End Sub
Public Sub Debug_ToExcel(Optional sheet As String = "OUTPUT")
    OutputBuffer_ToExcel obDebug, sheet
End Sub
Public Sub Debug_ToFile(filePath As String)
    OutputBuffer_ToFile obDebug, filePath
End Sub
Public Sub Debug_CM_Start(Optional numColumns As Integer = 2)
    OutputBuffer_CM_Start obDebug, numColumns
    
    OutputBuffer_CM_HBar obDebug ' Auto-add horizontal bar
End Sub
Public Sub Debug_CM_Print(colNum As Integer, str As String)
    OutputBuffer_CM_Print obDebug, colNum, str
End Sub
Public Sub Debug_CM_PrintRow(ParamArray outputArr() As Variant)
    OutputBuffer_CM_PrintRow obDebug, outputArr
End Sub
Public Sub Debug_CM_HBar()
    OutputBuffer_CM_HBar obDebug
End Sub
Public Sub Debug_CM_EndFlush(Optional minColWidth As Integer = 0, Optional colSep As String = "|", Optional linePrefix As String = "|", Optional lineSuffix As String = "|")
    
    
    OutputBuffer_CM_HBar obDebug ' Auto-add horizontal bar to end
    
    Dim obLineFrom As Long, obLineTo As Long, i As Long
    
    obLineFrom = obDebug.CurLine + 1
    OutputBuffer_CM_EndFlush obDebug, minColWidth, colSep, linePrefix, lineSuffix
    obLineTo = obDebug.CurLine
    
    If obLineTo - obLineFrom > 0 Then
        For i = obLineFrom To obLineTo
            Debug.Print obDebug.Lines(i) ' Print to immediate window
        Next i
    End If
End Sub
' Example: Debug_OutputVariableList("Var1", "s", "Var2", 4)
' Returns string: Var1='s', Var2=4
Public Function Debug_VarListString(ParamArray variablesAndValues() As Variant) As String
    
    Dim outStr As String, valStr As String, varName As String, val As Variant
    Dim valTypeIsStringLike As Boolean, valType As String
    
    Dim i As Long
    
    For i = LBound(variablesAndValues) To UBound(variablesAndValues) Step 2
        varName = variablesAndValues(i)
        valStr = ""
        
        If i + 1 <= UBound(variablesAndValues) Then
            val = variablesAndValues(i + 1)
            
            valType = VBA.typeName(val)
            valTypeIsStringLike = False
            
            
            Select Case valType
                Case "Boolean":
                    valStr = IIf(CBool(val), "True", "False")
                Case "String"
                    valStr = CStr(val)
                    valTypeIsStringLike = True
                Case Else:
                    If IsObject(val) Then
                        valStr = "<Object>"
                    ElseIf IsObject(val) Then
                        valStr = "<Array>"
                    Else
                        valStr = CStr(val)
                    End If
            End Select
            
            
            
            If valTypeIsStringLike Then
                valStr = Replace$(val, vbCrLf, "\n")
                valStr = Replace$(val, vbCr, "\n")
                valStr = Replace$(val, vbLf, "\n")
                valStr = "'" & valStr & "'"
            End If
        End If
        
        outStr = outStr & ", " & varName & "=" & valStr
    Next i
    
    outStr = Mid$(outStr, Len(", ") + 1)
    
    Debug_VarListString = outStr
    
End Function

