Attribute VB_Name = "JH_DebugHelper"
Option Explicit
Option Private Module




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

    
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheet)
    If Err.Number <> 0 Then
        ws.Cells.Clear
    Else
        Set ws = ThisWorkbook.Sheets.Add()
        ws.Name = sheet
    End If
    On Error GoTo 0
    
    
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
    
    
    
    'If Worksheet_SheetExists(sheet) Then
        Set ws = ThisWorkbook.Worksheets(sheet)
        ws.Cells.Clear
    'Else
    '    Set ws = ThisWorkbook.Sheets.Add()
    '    ws.Name = sheet
    'End If
    
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

