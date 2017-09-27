Attribute VB_Name = "SpreadsheetMapper"
Option Explicit

' Utility module to help map 1D worksheet data to data structures with parent/child relationships (2 levels)

Type SPREADSHEET_MAP_1L_PARENT
    PARENT_ID As String
    
    PARENT_MAP_TO_ROW As Integer
End Type

Type SPREADSHEET_MAP_1L
    PARENT_COUNT As Integer
    PARENT_MAP() As SPREADSHEET_MAP_1L_PARENT

    MAP_ROW_TO_PARENT() As Integer
    
    ROW_COUNT As Integer
End Type

Type SPREADSHEET_MAP_2L_PARENT
    PARENT_ID As String
    
    PARENT_MAP_TO_ROW As Integer
            
    CHILD_COUNT As Integer
    MAP_CHILD_TO_ROW() As Integer
End Type

Type SPREADSHEET_MAP_2L
    PARENT_COUNT As Integer
    PARENT_MAP() As SPREADSHEET_MAP_2L_PARENT

    MAP_ROW_TO_PARENT() As Integer
    MAP_ROW_TO_CHILD() As Integer
    
    ROW_COUNT As Integer
End Type




' Spreadsheet to Data Structure Mapping - 2 Dimensional
' Creates a mapping between a spreadsheet table to a 2D VBA data structure (an array of parent items, each parent item has one or more secondary items)
Public Function SpreadsheetTableToMultiLevelMap_2D(tableRange As Range, countColumn As Integer, keyColumn1 As Variant, Optional headerRowSpan As Integer = 1) As SPREADSHEET_MAP_2L

    Dim i As Integer, j As Integer
    Dim idxParent As Integer, idxChild As Integer

    Dim curRow As Integer
    
    Dim ssMap As SPREADSHEET_MAP_2L
    
    
   ' i = 2
    'rowCount = 0
    'Do While tableRange.Cells(i, keyColumn1)
    'Loop
    Dim rowStart As Integer, rowEnd As Integer
    
    ssMap.ROW_COUNT = Application.WorksheetFunction.CountA(tableRange.Offset(headerRowSpan).Columns(countColumn))
     
     ' bug fix: 2.10.2
    If ssMap.ROW_COUNT = 0 Then
        ssMap.PARENT_COUNT = 0
        SpreadsheetTableToMultiLevelMap_2D = ssMap
        Exit Function
    End If
    
    rowStart = headerRowSpan + 1
    rowEnd = ssMap.ROW_COUNT + headerRowSpan
    
    
    
    ReDim ssMap.MAP_ROW_TO_PARENT(rowStart To rowEnd) As Integer
    ReDim ssMap.MAP_ROW_TO_CHILD(rowStart To rowEnd) As Integer
    
    
    ' Examine spreadsheet range and determine number primary items, secondary items, and create mapping between rows and PO objects
    For curRow = rowStart To rowEnd
        idxParent = 0
        
        If True Then 'PO_QUEUE_TABLE_Range.Cells(curRow, C_PO_NUM).Value = "" Then
            Dim primaryID As String
            
            primaryID = ""
            
            ' key is multi-column?
            If IsArray(keyColumn1) Then
                For i = LBound(keyColumn1) To UBound(keyColumn1)
                    primaryID = primaryID & "|" & tableRange.Cells(curRow, keyColumn1(i)).Value
                Next i
                
                primaryID = Mid(primaryID, 2)
            Else
                primaryID = tableRange.Cells(curRow, keyColumn1).Value
            End If
            
            
        
            For i = 1 To ssMap.PARENT_COUNT
                If ssMap.PARENT_MAP(i).PARENT_ID = primaryID Then
                    idxParent = i
                    Exit For
                End If
            Next i
            
            If idxParent = 0 Then
                ssMap.PARENT_COUNT = ssMap.PARENT_COUNT + 1
                idxParent = ssMap.PARENT_COUNT
                ReDim Preserve ssMap.PARENT_MAP(1 To ssMap.PARENT_COUNT) As SPREADSHEET_MAP_2L_PARENT
                
                ssMap.PARENT_MAP(idxParent).PARENT_ID = primaryID
                ssMap.PARENT_MAP(idxParent).PARENT_MAP_TO_ROW = curRow
            End If
            
            
            ssMap.PARENT_MAP(idxParent).CHILD_COUNT = ssMap.PARENT_MAP(idxParent).CHILD_COUNT + 1
            
            
            ssMap.MAP_ROW_TO_PARENT(curRow) = idxParent
            ssMap.MAP_ROW_TO_CHILD(curRow) = ssMap.PARENT_MAP(idxParent).CHILD_COUNT
        End If
    Next curRow
    
    If ssMap.PARENT_COUNT = 0 Then
        SpreadsheetTableToMultiLevelMap_2D = ssMap
        Exit Function
    End If
    
    ' Create secondary item mapping

    For curRow = rowStart To rowEnd
        idxParent = ssMap.MAP_ROW_TO_PARENT(curRow)
        
           
        If idxParent > 0 Then
           idxChild = ssMap.MAP_ROW_TO_CHILD(curRow)
           
           If idxChild = 1 Then
               ReDim ssMap.PARENT_MAP(idxParent).MAP_CHILD_TO_ROW(1 To ssMap.PARENT_MAP(idxParent).CHILD_COUNT) As Integer
           End If
           
           ssMap.PARENT_MAP(idxParent).MAP_CHILD_TO_ROW(idxChild) = curRow

        End If
    Next curRow
    
    SpreadsheetTableToMultiLevelMap_2D = ssMap

End Function
' Spreadsheet to Data Structure Mapping - 21 Dimensional
' Creates a mapping between a spreadsheet table to a 2D VBA data structure (an array of primary items, each primary item has no secondary items
Public Function SpreadsheetTableToMultiLevelMap_1D(tableRange As Range, countColumn As Integer, keyColumn1 As Integer, Optional headerRowSpan As Integer = 1) As SPREADSHEET_MAP_1L

    Dim i As Integer, j As Integer
    Dim idxParent As Integer

    Dim curRow As Integer
    
    Dim ssMap As SPREADSHEET_MAP_1L
    
    
    Dim rowStart As Integer, rowEnd As Integer
    
    'Dim offsetRange As Range
    'Set offsetRange = tableRange
    'If headerRowSpan > 0 Then Set offsetRange = tableRange.Offset(headerRowSpan)
    
    ssMap.ROW_COUNT = Application.WorksheetFunction.CountA(tableRange.Offset(headerRowSpan).Columns(countColumn))
    
    rowStart = headerRowSpan + 1
    rowEnd = ssMap.ROW_COUNT + headerRowSpan
    
    
    
    ReDim ssMap.MAP_ROW_TO_PARENT(rowStart To rowEnd) As Integer
    'ReDim ssMap.MAP_ROW_TO_CHILD(rowStart To rowEnd) As Integer
    
    ' Examine spreadsheet range and determine number primary items, secondary items, and create mapping between rows and PO objects
    For curRow = rowStart To rowEnd
        idxParent = 0
        
        If True Then 'PO_QUEUE_TABLE_Range.Cells(curRow, C_PO_NUM).Value = "" Then
            'For i = 1 To ssMap.PARENT_COUNT
            '    If ssMap.PARENT_MAP(i).PARENT_ID = tableRange.Cells(curRow, keyColumn1).Value Then
            '        idxParent = i
            '        Exit For
            '    End If
            'Next i
            
            'If idxParent = 0 Then
                ssMap.PARENT_COUNT = ssMap.PARENT_COUNT + 1
                idxParent = ssMap.PARENT_COUNT
                ReDim Preserve ssMap.PARENT_MAP(1 To ssMap.PARENT_COUNT) As SPREADSHEET_MAP_1L_PARENT
                
                ssMap.PARENT_MAP(idxParent).PARENT_ID = tableRange.Cells(curRow, keyColumn1).Value
                ssMap.PARENT_MAP(idxParent).PARENT_MAP_TO_ROW = curRow
            'End If
            
            
            'ssMap.PARENT_MAP(idxParent).CHILD_COUNT = ssMap.PARENT_MAP(idxParent).CHILD_COUNT + 1
            
            
            ssMap.MAP_ROW_TO_PARENT(curRow) = idxParent
            'ssMap.MAP_ROW_TO_CHILD(curRow) = ssMap.PARENT_MAP(idxParent).CHILD_COUNT
        End If
    Next curRow
    
    
    SpreadsheetTableToMultiLevelMap_1D = ssMap

End Function
