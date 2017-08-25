Attribute VB_Name = "JH_PrePost"
Option Explicit


' JH_PrePost
' ------------------------------------------------------------------------------------------------------------------'
' VBA functions which support the Pre/Post analysis
'
'
' Joseph Huntley (joseph.huntley@vzw.com)
' ------------------------------------------------------------------------------------------------------------------




' Calculates a Pre or Post KPI value by the following process.
'     (1) Calculate the daily sum for all eNB/Sector/Carriers (can be a single ENB/S/C)
'     (2) Calculate the average across all days for the given percentile bounds (e.g. ignore the top x% and/or bottom y%)
Public Function PrePost_LTE_CalcKpiColumn_Aggr(dataRng As Range, kpiColumn As String, subject_enodeID As Variant, subject_sector As Variant, subject_carrier As Variant, refDate As Date, beforeDate As Boolean, Optional pcntileBoundStart As Double = 0, Optional pcntileBoundEnd As Double = 1) As Double

    'Dim subject_euTrancell As Variant
    
    'Dim dataRng As Range: Set dataRng = Range("ALPT_DT!A:BN")
    'Dim kpiColumn As String: kpiColumn = "DL_RLC_Layer_MByte"
    'Dim refDate As Date: refDate = #11/2/2014#
    'Dim subject_enodeID As Variant: Set subject_enodeID = Range("'LTE Results'!B11:B50")
    'Dim subject_sector As Variant: Set subject_sector = Range("'LTE Results'!C11:C50")
    'Dim subject_carrier As Variant: Set subject_carrier = Range("'LTE Results'!D11:D50")
    'Dim beforeDate As Boolean: beforeDate = True
    'Dim pcntileBoundStart As Single: pcntileBoundStart = 0
    'Dim pcntileBoundEnd As Single: pcntileBoundEnd = 1
    

    
    Dim i As Long, j As Long
    
    Dim subjectClusterList() As String
    Dim subjectClusterCount As Long
    
    If TypeName(subject_enodeID) = "Range" Then
        Debug.Assert TypeName(subject_sector) = "Range"
        Debug.Assert TypeName(subject_carrier) = "Range"
        
        Dim rngSubject_enodeID As Range, rngSubject_sector As Range, rngSubject_carrier As Range
    
        Set rngSubject_enodeID = subject_enodeID
        Set rngSubject_sector = subject_sector
        Set rngSubject_carrier = subject_carrier
        
        If rngSubject_enodeID.Rows.count <> rngSubject_sector.Rows.count Then Err.Raise -1, , "Number of items in subject sector must match number of eNBs"
        If rngSubject_enodeID.Rows.count <> rngSubject_carrier.Rows.count Then Err.Raise -1, , "Number of items in subject sector must match number of eNBs"
        
        subjectClusterCount = 0
                
        For i = 1 To rngSubject_enodeID.Rows.count
        
            If rngSubject_enodeID.Cells(i, 1) <> "" And rngSubject_sector.Cells(i, 1) <> "" And rngSubject_carrier.Cells(i, 1) <> "" Then
                subjectClusterCount = subjectClusterCount + 1
                
                ReDim Preserve subjectClusterList(1 To subjectClusterCount) As String
                
                subjectClusterList(subjectClusterCount) = rngSubject_enodeID.Cells(i, 1) & "-" & rngSubject_sector.Cells(i, 1) & "-" & rngSubject_carrier.Cells(i, 1)
            End If
        Next i
        
        Set rngSubject_enodeID = Nothing
        Set rngSubject_sector = Nothing
        Set rngSubject_carrier = Nothing
    End If
    
    
    
    Dim data() As Variant, numEntries As Long
    
    data = Worksheet_LoadColumnDataToArray(dataRng, Array("DAY", "ENODEB", "EUTRANCELL", "CARRIER", kpiColumn))
    
    
    Dim lbCols As Long, ubCols As Long
    Dim lbDataRows As Long, ubDataRows As Long
    Dim dpDate As Date, dpEnodeID As Long, dpSector As Integer, dpCarrier As Integer, dpKpiColData As Double
    
    Dim isInSubjectClusterList As Boolean, clusterItem As String
    
    Dim numDataPoints As Long
    Dim kpiDataPoints() As Double
    
    Dim dailyAggr_Count As Long
    Dim dailyAggr_Date() As Date
    Dim dailyAggr_kpiValue() As Double
    Dim dailyAggr_dataPointCount() As Long
    
    Dim dailyAggrIdx As Long
    
    dailyAggr_Count = 0
    


    lbCols = LBound(data, 1)
    ubCols = LBound(data, 1)
    lbDataRows = LBound(data, 2)
    ubDataRows = UBound(data, 2)
    numDataPoints = 0


    For i = lbDataRows To ubDataRows
        dpDate = data(lbCols + 0, i)
        dpEnodeID = data(lbCols + 1, i)
        dpSector = data(lbCols + 2, i)
        dpCarrier = data(lbCols + 3, i)
        dpKpiColData = data(lbCols + 4, i)
        
        
        If (dpDate < refDate And beforeDate = True) Or (dpDate > refDate And beforeDate = False) Then
            isInSubjectClusterList = False
            clusterItem = dpEnodeID & "-" & dpSector & "-" & dpCarrier
            
            For j = 1 To subjectClusterCount
                If subjectClusterList(j) = clusterItem Then
                    isInSubjectClusterList = True
                    Exit For
                End If
            Next j
        
            If isInSubjectClusterList Then
                'numDataPoints = numDataPoints + 1
                'ReDim Preserve kpiDataPoints(1 To numDataPoints) As Double
                'kpiDataPoints(numDataPoints) = dpKpiColData
                
                ' Is this day already in daily totals?
                dailyAggrIdx = -1
                
                For j = 1 To dailyAggr_Count
                    If dpDate = dailyAggr_Date(j) Then
                        dailyAggrIdx = j
                        Exit For
                    End If
                Next j
                
                ' -> NO. Then add it
                If dailyAggrIdx = -1 Then
                    dailyAggr_Count = dailyAggr_Count + 1
                    dailyAggrIdx = dailyAggr_Count
                    
                    ReDim Preserve dailyAggr_Date(1 To dailyAggr_Count) As Date
                    ReDim Preserve dailyAggr_kpiValue(1 To dailyAggr_Count) As Double
                    ReDim Preserve dailyAggr_dataPointCount(1 To dailyAggr_Count) As Long
                    
                    dailyAggr_Date(dailyAggr_Count) = dpDate
                    dailyAggr_kpiValue(dailyAggrIdx) = 0
                    dailyAggr_dataPointCount(dailyAggrIdx) = 0
                End If
                
                ' Add data point to daily totals
                dailyAggr_kpiValue(dailyAggrIdx) = dailyAggr_kpiValue(dailyAggrIdx) + dpKpiColData
                dailyAggr_dataPointCount(dailyAggrIdx) = dailyAggr_dataPointCount(dailyAggrIdx) + 1
            End If
            
        End If
    Next i
    
    
    PrePost_LTE_CalcKpiColumn_Aggr = PercentileAverage(dailyAggr_kpiValue, pcntileBoundStart, pcntileBoundEnd)


End Function

Public Function PrePost_LTE_CalcKpiColumn(dataRng As Range, kpiColumn As String, subject_enodeID As Long, subject_euTrancell As Integer, subject_carrier As Integer, refDate As Date, beforeDate As Boolean, Optional pcntileBoundStart As Double = 0, Optional pcntileBoundEnd As Double = 1) As Double

    'Dim dataRng As Range: Set dataRng = Range("ALPT_DT!A:BN")
    'Dim kpiColumn As String: kpiColumn = "DL_RLC_Layer_Mbyte"
    'Dim refDate As Date: refDate = #11/29/2014#
    'Dim subject_enodeID As Long: subject_enodeID = 100007
    'Dim subject_euTrancell As Integer: subject_euTrancell = 1
    'Dim subject_carrier As Integer: subject_carrier = 1
    'Dim beforeDate As Boolean: beforeDate = True
    'Dim pcntileBoundStart As Single: pcntileBoundStart = 0
    'Dim pcntileBoundEnd As Single: pcntileBoundEnd = 1
    
    
    Dim data() As Variant, numEntries As Long
    
    data = Worksheet_LoadColumnDataToArray(dataRng, Array("DAY", "ENODEB", "EUTRANCELL", "CARRIER", kpiColumn))
    
    
    Dim i As Long, j As Long
    Dim lbCols As Long, ubCols As Long
    Dim lbDataRows As Long, ubDataRows As Long
    Dim dpDate As Date, dpEnodeID As Long, dpEuTrancell As Integer, dpCarrier As Integer, dpKpiColData As Double
    
    Dim numDataPoints As Long
    Dim kpiDataPoints() As Double

    lbCols = LBound(data, 1)
    ubCols = LBound(data, 1)
    lbDataRows = LBound(data, 2)
    ubDataRows = UBound(data, 2)
    numDataPoints = 0


    For i = lbDataRows To ubDataRows
        dpDate = data(lbCols + 0, i)
        dpEnodeID = data(lbCols + 1, i)
        dpEuTrancell = data(lbCols + 2, i)
        dpCarrier = data(lbCols + 3, i)
        dpKpiColData = data(lbCols + 4, i)
        
        If dpKpiColData > 0 Then
            If (dpDate < refDate And beforeDate = True) Or (dpDate > refDate And beforeDate = False) Then
                If dpEnodeID = subject_enodeID And dpEuTrancell = subject_euTrancell And dpCarrier = subject_carrier Then
                    numDataPoints = numDataPoints + 1
                    ReDim Preserve kpiDataPoints(1 To numDataPoints) As Double
                    
                    kpiDataPoints(numDataPoints) = dpKpiColData
                End If
            End If
        End If
    Next i
    
    If kpiColumn = "RRC_Conn_Fail_den" And subject_enodeID = 98327 And subject_euTrancell = 2 And subject_carrier = 1 And beforeDate = False Then
        Debug.Print
    End If
    
    
    PrePost_LTE_CalcKpiColumn = PercentileAverage(kpiDataPoints, pcntileBoundStart, pcntileBoundEnd)

End Function


Public Function PrePost_DO_CalcEHRPD(dataRng As Range, perfTool As String, subject_SN As Long, subject_cell As Long, subject_sector As Integer, refDate As Date, beforeDate As Boolean, Optional pcntileBoundStart As Double = 0, Optional pcntileBoundEnd As Double = 1) As Double

    
    'Dim dataRng As Range: Set dataRng = Range("RTT!A:AN")
    'Dim perfTool As String: perfTool = "RTT"
    'Dim refDate As Date: refDate = #12/12/2014#
    'Dim subject_SN As Long: subject_SN = 54
    'Dim subject_cell As Integer: subject_cell = 31
    'Dim subject_sector As Integer: subject_sector = 1
    'Dim beforeDate As Boolean: beforeDate = False
    'Dim pcntileBoundStart As Single: pcntileBoundStart = 0
    'Dim pcntileBoundEnd As Single: pcntileBoundEnd = 1
    
    ' Calculates eHPRD KPI as the average of the daily totals, keeping only values which falls in the specified percentile range
    ' - Ignores all hours except hours 15,16,17 (HQ Report card metric)
    
    'beforeDate = False


    If perfTool <> "MPT" And perfTool <> "RTT" Then
        Err.Raise -1, , "Invalid performance tool: " & perfTool
    End If
    
    
    Dim kpiColumn As String
    
    Dim data() As Variant, numEntries As Long
    
    If perfTool = "RTT" Then
        kpiColumn = "eHRPD Conn #"
        data = Worksheet_LoadColumnDataToArray(dataRng, Array("Date", "Hour", "SN", "Cell", "Sector", kpiColumn))
    ElseIf perfTool = "MPT" Then
        kpiColumn = "EHRPD Access Count"
        data = Worksheet_LoadColumnDataToArray(dataRng, Array("Date", "Hr", "BSC", "BTS", "Sector", kpiColumn))
    End If
    
    
    
    Dim i As Long, j As Long
    Dim lbCols As Long, ubCols As Long
    Dim lbDataRows As Long, ubDataRows As Long
    Dim dpDate As Date, dpHr As Integer, dpSN As Integer, dpCell As Long, dpSector As Integer, dpKpiColData As Double
    
    Dim numDataPoints As Long
    Dim kpiDataPoints() As Double
    
    Dim dailyAggr_Count As Long
    Dim dailyAggr_Date() As Date
    Dim dailyAggr_Sum() As Double
    Dim dailyAggr_dataPointCount() As Long
    
    Dim dailyAggrIdx As Long
    
    dailyAggr_Count = 0
    

    lbCols = LBound(data, 1)
    ubCols = LBound(data, 1)
    lbDataRows = LBound(data, 2)
    ubDataRows = UBound(data, 2)


    For i = lbDataRows To ubDataRows
        dpDate = data(lbCols + 0, i)
        dpHr = data(lbCols + 1, i)
        dpSN = data(lbCols + 2, i)
        dpCell = data(lbCols + 3, i)
        dpSector = data(lbCols + 4, i)
        dpKpiColData = data(lbCols + 5, i)
        

        
        
        If (dpDate < refDate And beforeDate = True) Or (dpDate > refDate And beforeDate = False) Then
            If dpHr >= 15 And dpHr <= 17 Then
                If dpSN = subject_SN And dpCell = subject_cell And dpSector = subject_sector Then
                    ' Is this day already in daily totals?
                    dailyAggrIdx = -1
                    
                    For j = 1 To dailyAggr_Count
                        If dpDate = dailyAggr_Date(j) Then
                            dailyAggrIdx = j
                            Exit For
                        End If
                    Next j
                    
                    ' -> NO. Then add it
                    If dailyAggrIdx = -1 Then
                        dailyAggr_Count = dailyAggr_Count + 1
                        dailyAggrIdx = dailyAggr_Count
                        
                        ReDim Preserve dailyAggr_Date(1 To dailyAggr_Count) As Date
                        ReDim Preserve dailyAggr_Sum(1 To dailyAggr_Count) As Double
                        ReDim Preserve dailyAggr_dataPointCount(1 To dailyAggr_Count) As Long
                        
                        dailyAggr_Date(dailyAggr_Count) = dpDate
                        dailyAggr_Sum(dailyAggrIdx) = 0
                        dailyAggr_dataPointCount(dailyAggrIdx) = 0
                    End If
                    
                    ' Add eHRPD connections to daily totals
                    dailyAggr_Sum(dailyAggrIdx) = dailyAggr_Sum(dailyAggrIdx) + dpKpiColData
                    dailyAggr_dataPointCount(dailyAggrIdx) = dailyAggr_dataPointCount(dailyAggrIdx) + 1
                    
                End If
            End If
        End If
    Next i
    
    
    
    PrePost_DO_CalcEHRPD = PercentileAverage(dailyAggr_Sum, pcntileBoundStart, pcntileBoundEnd)
    

End Function
' Calculates a Pre or Post KPI value by the following process.
'     (1) Calculate the daily sum for all eNB/Sector/Carriers (can be a single ENB/S/C)
'     (2) Calculate the average across all days for the given percentile bounds (e.g. ignore the top x% and/or bottom y%)
Public Function PrePost_DO_CalcEHRPD_Aggr(dataRng As Range, perfTool As Variant, subject_SN As Variant, subject_cell As Variant, subject_sector As Variant, refDate As Date, beforeDate As Boolean, Optional pcntileBoundStart As Double = 0, Optional pcntileBoundEnd As Double = 1) As Double
                    'dataRng As Range, kpiColumn As String, subject_enodeID As Variant, subject_sector As Variant, subject_carrier As Variant, refDate As Date, beforeDate As Boolean, Optional pcntileBoundStart As Double = 0, Optional pcntileBoundEnd As Double = 1) As Double

    'Dim subject_euTrancell As Variant
    
    'Dim dataRng As Range: Set dataRng = Range("ALPT_DT!A:BN")
    'Dim kpiColumn As String: kpiColumn = "DL_RLC_Layer_MByte"
    'Dim refDate As Date: refDate = #11/2/2014#
    'Dim subject_enodeID As Variant: Set subject_enodeID = Range("'LTE Results'!B11:B50")
    'Dim subject_sector As Variant: Set subject_sector = Range("'LTE Results'!C11:C50")
    'Dim subject_carrier As Variant: Set subject_carrier = Range("'LTE Results'!D11:D50")
    'Dim beforeDate As Boolean: beforeDate = True
    'Dim pcntileBoundStart As Single: pcntileBoundStart = 0
    'Dim pcntileBoundEnd As Single: pcntileBoundEnd = 1
    


    If perfTool <> "MPT" And perfTool <> "RTT" Then
        Err.Raise -1, , "Invalid performance tool: " & perfTool
    End If
    
    
    
    Dim i As Long, j As Long
    
    Dim subjectClusterList() As String
    Dim subjectClusterCount As Long
    
    If TypeName(subject_cell) = "Range" Then
        Debug.Assert TypeName(subject_SN) = "Range"
        Debug.Assert TypeName(subject_sector) = "Range"
        
        Dim rngSubject_SN As Range, rngSubject_cell As Range, rngSubject_sector As Range
    
        Set rngSubject_SN = subject_SN
        Set rngSubject_cell = subject_cell
        Set rngSubject_sector = subject_sector
        
        If rngSubject_cell.Rows.count <> rngSubject_SN.Rows.count Then Err.Raise -1, , "Number of items in subject SN must match number of cells"
        If rngSubject_cell.Rows.count <> rngSubject_sector.Rows.count Then Err.Raise -1, , "Number of items in subject sector must match number of cells"
        
        subjectClusterCount = 0
                
        For i = 1 To rngSubject_cell.Rows.count
        
            If rngSubject_SN.Cells(i, 1) <> "" And rngSubject_cell.Cells(i, 1) <> "" And rngSubject_sector.Cells(i, 1) <> "" Then
                subjectClusterCount = subjectClusterCount + 1
                
                ReDim Preserve subjectClusterList(1 To subjectClusterCount) As String
                
                subjectClusterList(subjectClusterCount) = rngSubject_SN.Cells(i, 1) & "-" & rngSubject_cell.Cells(i, 1) & "-" & rngSubject_sector.Cells(i, 1)
            End If
        Next i
        
        Set rngSubject_SN = Nothing
        Set rngSubject_cell = Nothing
        Set rngSubject_sector = Nothing
    End If
    
    
    
    Dim kpiColumn As String
    
    Dim data() As Variant, numEntries As Long
    
    If perfTool = "RTT" Then
        kpiColumn = "eHRPD Conn #"
        data = Worksheet_LoadColumnDataToArray(dataRng, Array("Date", "Hour", "SN", "Cell", "Sector", kpiColumn))
    ElseIf perfTool = "MPT" Then
        kpiColumn = "EHRPD Access Count"
        data = Worksheet_LoadColumnDataToArray(dataRng, Array("Date", "Hr", "BSC", "BTS", "Sector", kpiColumn))
    End If
    
    
    
    Dim lbCols As Long, ubCols As Long
    Dim lbDataRows As Long, ubDataRows As Long
    Dim dpDate As Date, dpHr As Integer, dpSN As Integer, dpCell As Long, dpSector As Integer, dpKpiColData As Double
    
    Dim isInSubjectClusterList As Boolean, clusterItem As String
    
    Dim numDataPoints As Long
    Dim kpiDataPoints() As Double
    
    Dim dailyAggr_Count As Long
    Dim dailyAggr_Date() As Date
    Dim dailyAggr_kpiValue() As Double
    Dim dailyAggr_dataPointCount() As Long
    
    Dim dailyAggrIdx As Long
    
    dailyAggr_Count = 0
    


    lbCols = LBound(data, 1)
    ubCols = LBound(data, 1)
    lbDataRows = LBound(data, 2)
    ubDataRows = UBound(data, 2)
    numDataPoints = 0


    For i = lbDataRows To ubDataRows
        dpDate = data(lbCols + 0, i)
        dpHr = data(lbCols + 1, i)
        dpSN = data(lbCols + 2, i)
        dpCell = data(lbCols + 3, i)
        dpSector = data(lbCols + 4, i)
        dpKpiColData = data(lbCols + 5, i)
        
        
        If (dpDate < refDate And beforeDate = True) Or (dpDate > refDate And beforeDate = False) Then
            If dpHr >= 15 And dpHr <= 17 Then
                isInSubjectClusterList = False
                clusterItem = dpSN & "-" & dpCell & "-" & dpSector
                
                For j = 1 To subjectClusterCount
                    If subjectClusterList(j) = clusterItem Then
                        isInSubjectClusterList = True
                        Exit For
                    End If
                Next j
            
                If isInSubjectClusterList Then
                    'numDataPoints = numDataPoints + 1
                    'ReDim Preserve kpiDataPoints(1 To numDataPoints) As Double
                    'kpiDataPoints(numDataPoints) = dpKpiColData
                    
                    ' Is this day already in daily totals?
                    dailyAggrIdx = -1
                    
                    For j = 1 To dailyAggr_Count
                        If dpDate = dailyAggr_Date(j) Then
                            dailyAggrIdx = j
                            Exit For
                        End If
                    Next j
                    
                    ' -> NO. Then add it
                    If dailyAggrIdx = -1 Then
                        dailyAggr_Count = dailyAggr_Count + 1
                        dailyAggrIdx = dailyAggr_Count
                        
                        ReDim Preserve dailyAggr_Date(1 To dailyAggr_Count) As Date
                        ReDim Preserve dailyAggr_kpiValue(1 To dailyAggr_Count) As Double
                        ReDim Preserve dailyAggr_dataPointCount(1 To dailyAggr_Count) As Long
                        
                        dailyAggr_Date(dailyAggr_Count) = dpDate
                        dailyAggr_kpiValue(dailyAggrIdx) = 0
                        dailyAggr_dataPointCount(dailyAggrIdx) = 0
                    End If
                    
                    ' Add data point to daily totals
                    dailyAggr_kpiValue(dailyAggrIdx) = dailyAggr_kpiValue(dailyAggrIdx) + dpKpiColData
                    dailyAggr_dataPointCount(dailyAggrIdx) = dailyAggr_dataPointCount(dailyAggrIdx) + 1
                End If
                
            End If
        End If
    Next i
    
    

    
    PrePost_DO_CalcEHRPD_Aggr = PercentileAverage(dailyAggr_kpiValue, pcntileBoundStart, pcntileBoundEnd)
    
    

End Function
Sub PrePost_Download_ALPT_Data()

    Dim i As Long
    
    Dim activationDate As Date, daysBeforeAndAfter As Integer
    Dim activatedEnodeID As Long
    Dim includeAwsOffload As Boolean
    Dim offloadClusterDef As String

    
    
    activatedEnodeID = Evaluate("=CFG_ENODEB_ID")
    daysBeforeAndAfter = Evaluate("=CFG_PREPOST_ACT_DAYS")
    activationDate = Evaluate("=CFG_ACTIVATION_DATE")
    offloadClusterDef = Evaluate("=TIER1_NEIGHBORS_LTE")
    includeAwsOffload = Evaluate("=CFG_AWS_OFFLOAD")


    Dim clusterDefList_LTE() As String, enodeList() As Long
    Dim daysToTrend As Integer, dateEnd As Date
    
    daysToTrend = 2 * daysBeforeAndAfter
    dateEnd = DateAdd("d", daysBeforeAndAfter, activationDate)
    
    If dateEnd > Now() Then
        MsgBox "Activation Date (" & activationDate & ") + " & daysBeforeAndAfter & " must before today's date."
        Exit Sub
    End If
    
    ' Add additional activations eNBs to data
    Dim rngActivatedEnodeID_List As Range
    Dim additionalActivatedEnode_Cluster As String
    
    Set rngActivatedEnodeID_List = ThisWorkbook.Names("CFG_ENODEB_ID_LIST").RefersToRange
    additionalActivatedEnode_Cluster = ""
    
    For i = 1 To rngActivatedEnodeID_List.Rows.count
        If rngActivatedEnodeID_List.Cells(i, 1) <> "" Then
            If rngActivatedEnodeID_List.Cells(i, 1) > 1000 Then ' Valid eNB ID?
                If rngActivatedEnodeID_List.Cells(i, 1) <> activatedEnodeID Then
                    additionalActivatedEnode_Cluster = additionalActivatedEnode_Cluster & rngActivatedEnodeID_List.Cells(i, 1) & "-1,2,3,4,5,6; "
                End If
            Else
                Err.Raise -1, , "Invalid eNB ID: " & rngActivatedEnodeID_List.Cells(i, 1)
            End If
        End If
    Next i
    

    offloadClusterDef = activatedEnodeID & "-1,2,3,4,5,6" & _
        IIf(Len(additionalActivatedEnode_Cluster) > 0, "; " & additionalActivatedEnode_Cluster, "") & _
        IIf(Len(offloadClusterDef) > 0, "; " & offloadClusterDef, "")
    
    clusterDefList_LTE = ClusterDef_Expand(offloadClusterDef)
        
    If includeAwsOffload Then clusterDefList_LTE = ClusterDef_LTE_Add_AWS_eNodeBs(clusterDefList_LTE)
    
    



    
    enodeList = ClusterDef_Extract_eNodeList(clusterDefList_LTE)
        
    
    

    Dim passNo As Integer, destSheet As String
    Dim reportType As String

    Dim myURL As String, queryData As String
    
    For passNo = 1 To 2
    
        destSheet = "ALPT_" & Choose(passNo, "DT", "BH_RRC")
        reportType = Choose(passNo, "Daily Totals", "euCell_RrcConn_BH")
        
    
        queryData = URL_BuildQueryString( _
            "action", "exectmpl", _
            "tmpl_rpt", "w503686|||PrePost_Macro_Exhaust", _
            "name", "PrePost_Macro_Exhaust_" & Format$(activatedEnodeID, "000000"), _
            "user", Environ$("username"), _
            "rpttype", reportType, _
            "edate", Format$(dateEnd, "yyyy-mm-dd"), _
            "num_days", daysToTrend, _
            "enodeb", Array_ToString(enodeList, "000000") _
        )
        
        'euCell_RrcConn_BH
        'Daily Totals
        'Hourly Totals
        
    
        myURL = "http://alpt.vh.eng.vzwcorp.com:8282/alte/reportWebService.htm?" & queryData
        
        
        Debug.Print myURL
        
        
        Excel_AppUpdates_Disable
        
    
        GetWebData_CSV myURL, destSheet

        
        
        ' Remove cell/sectors not in cluster list
        Dim colNbr_ENODEB As Integer
        Dim colNbr_EUTRANCELL As Integer
        
        
        Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(destSheet)
        colNbr_ENODEB = Worksheet_GetColumnByName(ws, "ENODEB")
        colNbr_EUTRANCELL = Worksheet_GetColumnByName(ws, "EUTRANCELL")
        
        Debug.Assert colNbr_ENODEB > 0
        Debug.Assert colNbr_EUTRANCELL > 0
        
        Dim clusterColumnNumbers As Variant: clusterColumnNumbers = Array(colNbr_ENODEB, colNbr_EUTRANCELL)
        Dim clusterColumnFormats As Variant:  clusterColumnFormats = Array("000000", "0")
        
        
        
        ClusterDef_FilterWorksheet ws, clusterDefList_LTE, clusterColumnNumbers, clusterColumnFormats


    
    
        Worksheet_FixNumbersStoredAsText ws
    

        Excel_AppUpdates_Restore
    
    Next passNo
    

End Sub
Sub PrePost_Download_MPT_Data()

    
    Dim activationDate As Date, daysBeforeAndAfter As Integer
    Dim activatedEnodeID As Long
    Dim offloadClusterDef As String
    
    
    activatedEnodeID = Evaluate("=CFG_ENODEB_ID")
    daysBeforeAndAfter = Evaluate("=CFG_PREPOST_ACT_DAYS")
    activationDate = Evaluate("=CFG_ACTIVATION_DATE")
    offloadClusterDef = Evaluate("=TIER1_NEIGHBORS_1XDO")


    Dim clusterDefList_LTE() As String, enodeList() As Long
    Dim daysToTrend As Integer, dateEnd As Date
    
    daysToTrend = 2 * daysBeforeAndAfter
    dateEnd = DateAdd("d", daysBeforeAndAfter, activationDate)
    
    If dateEnd > Now() Then
        MsgBox "Activation Date (" & activationDate & ") + " & daysBeforeAndAfter & " must before today's date."
        Exit Sub
    End If
    
    Dim clusterDefList_1XDO() As String
    
    clusterDefList_1XDO = ClusterDef_Expand(offloadClusterDef)
    
    

    
    Dim clusterDefList_1XDO_NoMarket() As String
    
    ' Strip market part from cluster - we can do this because cells in central PA markets do not conflict
    clusterDefList_1XDO_NoMarket = ClusterDef_RemoveMarketPart(clusterDefList_1XDO)
    
    
    MPT_CellGroupReport_Cluster destWorksheet:="MPT", techType:="DO", market:="centralpa", _
        clusterDef:=clusterDefList_1XDO_NoMarket, _
        reportType:="hourly", reportGroupBy:="sector", reportContent:="perf", _
        dateEnd:=dateEnd, numberOfDaysToTrend:=daysToTrend
        
        
    Exit Sub
    
    
    Dim colNbr_SysID As Integer
    Dim colNbr_SN As Integer
    Dim colNbr_BTS As Integer
    Dim colNbr_Sect As Integer
    
    
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("MPT")
    colNbr_BTS = Worksheet_GetColumnByName(ws, "BTS")
    colNbr_Sect = Worksheet_GetColumnByName(ws, "Sector")

    Debug.Assert colNbr_BTS > 0
    Debug.Assert colNbr_Sect > 0
    
    
    
    Dim clusterColumnNumbers As Variant: clusterColumnNumbers = Array(colNbr_BTS, colNbr_Sect)
    
    ClusterDef_FilterWorksheet ws, clusterDefList_1XDO_NoMarket, clusterColumnNumbers
    
    'Excel_AppUpdates_Disable
    
    
    
    
    'Excel_AppUpdates_Restore
    
End Sub



Private Sub PrePost_RefreshPivotTableData()

    
    ' Refresh Pivot table data
    Dim ws As Worksheet
    Dim pt As PivotTable
    
    
    
    Dim i As Integer
    
    On Error Resume Next
    
    Excel_AppUpdates_Disable
    
    
    For Each ws In ThisWorkbook.Sheets
        For Each pt In ws.PivotTables
            If pt.Name Like "ptPPD_*" Then
                pt.RefreshTable
                
                pt.Update
            End If
        Next pt
        
    
    Next ws
    
    
    Excel_AppUpdates_Restore
    
    On Error GoTo 0

End Sub
Private Sub PrePost_RefreshPivotCharts()

    Dim chtObj As Excel.ChartObject
    Dim cht As Excel.chart
    Dim chtSeries As Excel.Series
    Dim axisObj As Excel.Axis
    
   
    For Each chtObj In ThisWorkbook.Worksheets("Graphs").ChartObjects

        chtObj.chart.ChartType = xlLine
        
        Set axisObj = chtObj.chart.Axes(xlCategory)
        
        axisObj.TickLabelSpacing = 7 ' 7 days
    
      
    Next chtObj
    
    
    For Each chtObj In ThisWorkbook.Worksheets("Cluster Results").ChartObjects

        chtObj.chart.ChartType = xlLine
        
        Set axisObj = chtObj.chart.Axes(xlCategory)
        
        axisObj.TickLabelSpacing = 7 ' 7 days
    
      
    Next chtObj

End Sub

Private Sub PrePost_PrepareWorksheet_RTT()

    
    Dim ws As Worksheet
    Dim rowCount As Long
    
    Set ws = Sheets("RTT")
    
    
    
    
    
    Debug.Print
    
    Dim col As Integer
    Dim R As Range
    
    
    rowCount = Application.WorksheetFunction.CountA(ws.columns(1))


    Dim colNbr_SN As Integer
    Dim colNbr_CSC As Integer
    Dim colNbr_Cell As Integer
    Dim colNbr_Sector As Integer
    Dim colNbr_ECP_Cell_Sect As Integer
    Dim colNbr_Is_Surr_Cluster As Integer
    
    colNbr_SN = Worksheet_GetColumnByName(ws, "SN")
    colNbr_CSC = Worksheet_GetColumnByName(ws, "CSC")
    
    If colNbr_SN = 0 Or colNbr_CSC = 0 Then
        Exit Sub
    End If
    
    colNbr_Cell = Worksheet_GetColumnByName(ws, "Cell")
    
    If colNbr_Cell = 0 Then
        colNbr_Cell = colNbr_CSC + 1
        col = colNbr_Cell
        ws.columns(col).Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        ws.Cells(1, col) = "Cell"
        ws.Range(ws.Cells(2, col), ws.Cells(rowCount, col)).FormulaR1C1 = "=LEFT(RC" & colNbr_CSC & ",FIND(""."", RC" & colNbr_CSC & ")-1)"
    End If

    colNbr_Sector = Worksheet_GetColumnByName(ws, "Sector")
    
    If colNbr_Sector = 0 Then
        colNbr_Sector = colNbr_Cell + 1
        col = colNbr_Sector
        ws.columns(col).Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        ws.Cells(1, col) = "Sector"
        ws.Range(ws.Cells(2, col), ws.Cells(rowCount, col)).FormulaR1C1 = "=MID(RC" & colNbr_CSC & ", FIND(""."", RC" & colNbr_CSC & ")+1, FIND(""."", RC" & colNbr_CSC & ",FIND(""."", RC" & colNbr_CSC & ") + 1) - FIND(""."", RC" & colNbr_CSC & ") -1 )"
    End If
    
    'colNbr_ECP_Cell_Sect = colNbr_Sector + 1
    'col = colNbr_ECP_Cell_Sect
    'ws.Columns(col).Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    'ws.Cells(1, col) = "ECP_Cell_Sector"
    'ws.Range(ws.Cells(2, col), ws.Cells(rowCount, col)).FormulaR1C1 = "=INDEX(MTSO_TABLE_ECP, MATCH(RC" & colNbr_SN & ",MTSO_TABLE_SN,0)) & ""-"" & LEFT(RC" & colNbr_CSC & ",FIND(""."", RC" & colNbr_CSC & ")-1) & ""-"" & MID(RC" & colNbr_CSC & ", FIND(""."", RC" & colNbr_CSC & ")+1, FIND(""."", RC" & colNbr_CSC & ",FIND(""."", RC" & colNbr_CSC & ") + 1) - FIND(""."", RC" & colNbr_CSC & ") -1 )"



    ' Filter only hours 15-17
    
        
End Sub
Private Sub PrePost_PrepareWorksheet_MPT()

    
    Dim ws As Worksheet
    Dim rowCount As Long
    
    Set ws = Sheets("MPT")
    
    
       
    
    
    
    
    Dim col As Integer
    Dim R As Range
    
    
    rowCount = Application.WorksheetFunction.CountA(ws.columns(1))


    Dim colNbr_BTS As Integer
    Dim colNbr_Cell As Integer
    
    colNbr_BTS = Worksheet_GetColumnByName(ws, "BTS")
    
    If colNbr_BTS = 0 Then
        Exit Sub
    End If
    
    colNbr_Cell = Worksheet_GetColumnByName(ws, "Cell")
    
    If colNbr_Cell = 0 Then
        colNbr_Cell = colNbr_BTS + 1
        col = colNbr_Cell
        ws.columns(col).Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        ws.Cells(1, col) = "Cell"
        ws.Range(ws.Cells(2, col), ws.Cells(rowCount, col)).FormulaR1C1 = "=MOD(RC" & colNbr_BTS & ",1000)"
    End If

    

    ' Filter only hours 15-17
    
        
End Sub

Private Sub PrePost_PrepareWorksheet_ALPT(sheetName As String)

    
    Dim ws As Worksheet
    Dim rowCount As Long
    
    Set ws = Sheets(sheetName)
    
    Dim col As Integer
    Dim R As Range
    
    
    rowCount = Application.WorksheetFunction.CountA(ws.columns(1))


    Dim colNbr_SITE As Integer
    Dim colNbr_ENODEB As Integer
    Dim colNbr_EUTRANCELL As Integer
    Dim colNbr_CARRIER As Integer
    Dim colNbr_BAND As Integer
    Dim colNbr_Is_Subject_Enode As Integer
    Dim colNbr_Cell_Group As Integer
    
    colNbr_SITE = Worksheet_GetColumnByName(ws, "SITE")
    colNbr_ENODEB = Worksheet_GetColumnByName(ws, "ENODEB")
    colNbr_EUTRANCELL = Worksheet_GetColumnByName(ws, "EUTRANCELL")
    colNbr_CARRIER = Worksheet_GetColumnByName(ws, "CARRIER")
    
    If colNbr_SITE = 0 Or colNbr_ENODEB = 0 Or colNbr_EUTRANCELL = 0 Or colNbr_CARRIER = 0 Then
        Exit Sub
    End If
    
    
    colNbr_BAND = Worksheet_GetColumnByName(ws, "BAND")
    
    If colNbr_BAND = 0 Then
        colNbr_BAND = IIf(colNbr_CARRIER > 0, colNbr_CARRIER, colNbr_EUTRANCELL) + 1
        col = colNbr_BAND
        ws.columns(col).Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        ws.Cells(1, col) = "BAND"
        
        If colNbr_CARRIER > 0 Then ' Try to use carrier for determining band first. Otherwise. Use eNB ID
            ws.Range(ws.Cells(2, col), ws.Cells(rowCount, col)).FormulaR1C1 = "=IFERROR(CHOOSE(RC" & colNbr_CARRIER & ", ""700"", ""AWS"", ""AWS"", ""PCS-LTE""),"""")"
        Else
            ws.Range(ws.Cells(2, col), ws.Cells(rowCount, col)).FormulaR1C1 = "=IF(RC" & colNbr_ENODEB & ">=300000,""AWS"",""700"")"
        End If
    
    End If
    
    
    colNbr_Is_Subject_Enode = Worksheet_GetColumnByName(ws, "Is_Subj_Enode")
     
    If colNbr_Is_Subject_Enode = 0 Then
        colNbr_Is_Subject_Enode = colNbr_BAND + 1
        col = colNbr_Is_Subject_Enode
        ws.columns(col).Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        ws.Cells(1, col) = "Is_Subj_Enode"
        'ws.Range(ws.Cells(2, col), ws.Cells(rowCount, col)).FormulaR1C1 = "=IF(MOD(RC" & colNbr_ENODEB & ",300000)=CFG_ENODEB_ID,""YES"",""NO"")"
        ws.Range(ws.Cells(2, col), ws.Cells(rowCount, col)).FormulaR1C1 = "=IF(OR(MOD(RC" & colNbr_ENODEB & ",300000)=CFG_ENODEB_ID,NOT(ISNA(MATCH(MOD(RC" & colNbr_ENODEB & ",300000),CFG_ENODEB_ID_LIST,0)))),""YES"",""NO"")"
        
    End If
    
    
    colNbr_Cell_Group = Worksheet_GetColumnByName(ws, "Cell_Group")
     
    If colNbr_Cell_Group = 0 Then
        colNbr_Cell_Group = colNbr_Is_Subject_Enode + 1
        col = colNbr_Cell_Group
        ws.columns(col).Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        ws.Cells(1, col) = "Cell_Group"
        ws.Range(ws.Cells(2, col), ws.Cells(rowCount, col)).FormulaR1C1 = "=IF(RC" & colNbr_Is_Subject_Enode & "=""YES"",CFG_CLUSTER_NAME,""1st Tier Neighbors"")"
        
    End If
    
    
    
    If Worksheet_GetColumnByName(ws, "SITE_ENB_S_C") = 0 Then
        col = colNbr_BAND + 1
        ws.columns(col).Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        ws.Cells(1, col) = "SITE_ENB_S_C"
        ws.Range(ws.Cells(2, col), ws.Cells(rowCount, col)).FormulaR1C1 = "=CONCATENATE(RC" & colNbr_SITE & ", ""-"", RC" & colNbr_ENODEB & ", ""-"", RC" & colNbr_EUTRANCELL & ", ""-"", RC" & colNbr_CARRIER & ")"
    End If
    
    Dim resultCategoryTypes As Variant
    Dim i As Integer
    Dim colTemplate As String, colName As String, colNbr As Long
    
    resultCategoryTypes = Array("RRC Conn Atts", "RRC Drops", "User Throughput", "HO Fails", "RRC Drop %", "RRC Conn Fail %", "HO Fail %")
    
    col = colNbr_Is_Subject_Enode + 1
    
    For i = LBound(resultCategoryTypes) To UBound(resultCategoryTypes)
        colName = "LTE Result Category - " & resultCategoryTypes(i)
        colNbr = Worksheet_GetColumnByName(ws, colName)
        
        If colNbr = 0 Then
            colTemplate = "=IFERROR(OFFSET(LTE_RESULT_DESC_TABLE,MATCH(RC" & colNbr_ENODEB & " & ""-"" & RC" & colNbr_EUTRANCELL & " & ""-"" & RC" & colNbr_CARRIER & ",LTE_RESULT_UNIQUE_ID,0)-1,MATCH(""" & CStr(resultCategoryTypes(i)) & "*"",LTE_RESULT_DESC_TABLE_HEADERS,0)-1,1,1),"""")"
            'colTemplate = CStr(resultCategoryTypes(i))
        
            col = col + 1
            ws.columns(col).Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
            ws.Cells(1, col) = colName
            ws.Range(ws.Cells(2, col), ws.Cells(rowCount, col)).FormulaR1C1 = colTemplate

            colNbr = col
        End If
        
        
        colName = "Graph Category - " & resultCategoryTypes(i)
        
        If Worksheet_GetColumnByName(ws, colName) = 0 Then
            col = colNbr + 1
            ws.columns(col).Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
            ws.Cells(1, col) = colName
            ws.Range(ws.Cells(2, col), ws.Cells(rowCount, col)).FormulaR1C1 = "=IF(AND(RC" & colNbr_ENODEB & "=CFG_ENODEB_ID,RC" & colNbr_CARRIER & "=1),""Subject 700 ENB ("" & CHOOSE(RC" & colNbr_EUTRANCELL & ",""Alpha"",""Beta"",""Gamma"") & "")"",RC" & colNbr & ")"
    
        End If
        
    Next i


    '
End Sub


Public Sub PrePost_PrepareWorksheetsAndRefreshData()

    Excel_AppUpdates_Disable
    
    PrePost_PrepareWorksheet_ALPT "ALPT_DT"
    PrePost_PrepareWorksheet_ALPT "ALPT_BH_RRC"
    PrePost_PrepareWorksheet_RTT
    PrePost_PrepareWorksheet_MPT
    
    PrePost_RefreshPivotTableData
    

    
    PrePost_RefreshPivotCharts


    Excel_AppUpdates_Restore

End Sub


Public Sub PrePost_AutoCalcCounty()

    
    
    Dim tmpLat As Variant, tmpLon As Variant
    Dim sLat As Double, sLong As Double
    
    Dim test2 As Variant

    
    
    
    tmpLat = Evaluate(ThisWorkbook.Names("CFG_LOCATION_LAT").value)
    tmpLon = Evaluate(ThisWorkbook.Names("CFG_LOCATION_LON").value)
    
    If tmpLat = 0 Or tmpLat = "" Or tmpLon = 0 Or tmpLon = "" Then
        MsgBox "Latitude/Longitude is required"
        Exit Sub
    End If
    

    sLat = CDbl(tmpLat)
    sLong = CDbl(tmpLon)


    Dim geoCodeResult As Google_GeoCode_Reverse_Result
    
    geoCodeResult = Google_GeoCode_Reverse(sLat, sLong)
    
    
    Dim outRng As Range
    
    Set outRng = ThisWorkbook.Names("PARAM_COUNTY_OUT").RefersToRange


    If geoCodeResult.status_code = "OK" Then
        outRng.value = geoCodeResult.county & ", " & geoCodeResult.state_short
    Else
        outRng.value = "Error: " & geoCodeResult.status_description
    End If

End Sub
Public Sub PrePost_AutoCalculate1stTierNeighbors()


    

    Dim c  As Variant: c = Array(39.73875, -75.4541)

    ' North Percy: 39.962208, -75.153278
    ' Poplar St: 39.965778, -75.132778
    ' Monticello: 39.937361, -77.687464
    ' WIL PERSES: 39.73875, -75.4541
    
    
    Dim sLat As Double, sLong As Double
    Dim test2 As Variant

    
    Dim geoPlanDataFile As String, maxNeighborDistance As Double
    
    geoPlanDataFile = Evaluate("=RGNCFG_1TN_GeoPlanCellFilePath")
    maxNeighborDistance = Evaluate("=RGNCFG_1TN_MaxNeighborDistance")
    
    
    If Dir(geoPlanDataFile) = "" Then
        MsgBox "Cell file not found:  " & vbCrLf & vbCrLf & geoPlanDataFile
        Exit Sub
    End If
    
    
    
    Dim tmpLat As Variant, tmpLon As Variant
    
    tmpLat = Evaluate("=CFG_LOCATION_LAT")
    tmpLon = Evaluate("=CFG_LOCATION_LON")
    
    If tmpLat = 0 Or tmpLat = "" Or tmpLon = 0 Or tmpLon = "" Then
        MsgBox "Latitude/Longitude is required"
        Exit Sub
    End If
    

    sLat = CDbl(tmpLat)
    sLong = CDbl(tmpLon)

   
    Dim cellList_850() As CellSite
    Dim cellList_700() As CellSite


    cellList_850 = LoadCellListFromGeoPlanReport(geoPlanDataFile, "Cellular")
    cellList_700 = LoadCellListFromGeoPlanReport(geoPlanDataFile, "Upper 700 MHz")



    Dim neighbors_850() As NeighborCell
    Dim neighbors_700() As NeighborCell

    neighbors_850 = Calc1stTierNeighbors(cellList_850, sLat, sLong, maxNeighborDistance, True)
    neighbors_700 = Calc1stTierNeighbors(cellList_700, sLat, sLong, maxNeighborDistance, True)

    Debug.Print "-"

    Debug.Print "700: " & NeighborCells_GenerateClusterDefStr(neighbors_700)
    Debug.Print "850: " & NeighborCells_GenerateClusterDefStr(neighbors_850)



    ' ----------------------------------------------------------------------------------------
    ' Begin - Combine 700 and 850 neighbors into one list, joining together cells where appropiate
    ' ----------------------------------------------------------------------------------------
    Dim i As Long
    Dim cell_uid As String, arrIdx As Long

    Dim cells_Count As Long
    Dim cells_UIDs() As String
    Dim cells_CellNames() As String
    Dim cells_NeighborIdx_700() As Long
    Dim cells_NeighborIdx_850() As Long


    ' Begin - Create initial list of only 700 neighbors
    cells_Count = UBound(neighbors_700) - LBound(neighbors_700) + 1

    ReDim cells_UIDs(1 To cells_Count) As String
    ReDim cells_CellNames(1 To cells_Count) As String
    ReDim cells_NeighborIdx_700(1 To cells_Count) As Long
    ReDim cells_NeighborIdx_850(1 To cells_Count) As Long

    For i = LBound(neighbors_700) To UBound(neighbors_700)
        cell_uid = neighbors_700(i).cell.SwitchID & "-" & neighbors_700(i).cell.CellNum

        cells_UIDs(i - LBound(neighbors_700) + 1) = cell_uid
        cells_CellNames(i - LBound(neighbors_700) + 1) = neighbors_700(i).cell.Name
        cells_NeighborIdx_700(i - LBound(neighbors_700) + 1) = i
        cells_NeighborIdx_850(i - LBound(neighbors_700) + 1) = -1
    Next i
    ' End - Create initial list of only 700 neighbors


    ' Begin - Add 850 neighbors
    
    
    For i = LBound(neighbors_850) To UBound(neighbors_850)
        cell_uid = neighbors_850(i).cell.SwitchID & "-" & neighbors_850(i).cell.CellNum

        arrIdx = Array_Find(cells_UIDs, cell_uid)

        If arrIdx < 0 Then ' not found
            cells_Count = cells_Count + 1
    
            ReDim Preserve cells_UIDs(1 To cells_Count) As String
            ReDim Preserve cells_CellNames(1 To cells_Count) As String
            ReDim Preserve cells_NeighborIdx_700(1 To cells_Count) As Long
            ReDim Preserve cells_NeighborIdx_850(1 To cells_Count) As Long
    
            cells_UIDs(cells_Count) = cell_uid
            cells_CellNames(cells_Count) = neighbors_850(i).cell.Name
            cells_NeighborIdx_700(cells_Count) = -1
            cells_NeighborIdx_850(cells_Count) = i
        Else
            cells_NeighborIdx_850(arrIdx) = i
        End If
    Next i
  
    ' Begin - Add 850 neighbors

    Debug.Print
    ' ----------------------------------------------------------------------------------------
    ' End - Combine 700 and 850 neighbors into one list, joining together cells where appropiate
    ' ----------------------------------------------------------------------------------------


    Excel_AppUpdates_Disable
    
    Dim rng As Range
    
    Set rng = ThisWorkbook.Names("TIER1_NEIGHBORS_AUTOCALC_TABLE").RefersToRange
    
    rng.Cells.ClearContents
    
    For i = 1 To cells_Count
        rng.Cells(i, 1) = cells_CellNames(i)
        If cells_NeighborIdx_700(i) >= 0 Then rng.Cells(i, 2) = NeighborCell_GenerateClusterDefStr(neighbors_700(cells_NeighborIdx_700(i)))
        If cells_NeighborIdx_850(i) >= 0 Then rng.Cells(i, 3) = NeighborCell_GenerateClusterDefStr(neighbors_850(cells_NeighborIdx_850(i)))
    Next i
    
    
    Excel_AppUpdates_Restore


End Sub


Public Sub PrePost_AutoCalculate1stTierNeighbors_Cluster()

    
    Dim i As Long, j As Long, k As Long
    
    
    Dim geoPlanDataFile As String, maxNeighborDistance As Double
    
    geoPlanDataFile = Evaluate("=RGNCFG_1TN_GeoPlanCellFilePath") & "sss"
    maxNeighborDistance = Evaluate("=RGNCFG_1TN_MaxNeighborDistance")
    
    
    If Dir(geoPlanDataFile) = "" Then
        MsgBox "Cell file not found:  " & vbCrLf & vbCrLf & geoPlanDataFile
        Exit Sub
    End If
    
    
    
    
    
    Dim clusterBandClass As Variant, geoplanBandClass As String
    
    clusterBandClass = Evaluate("=CLUSTER_1TN_BAND_CLASS")
    
    If clusterBandClass <> 1 And clusterBandClass <> 2 Then '1=LTE, 2=1X/DO
        MsgBox "Invalid selected band class: select either LTE or 1X/DO"
        Exit Sub
    End If
    
    geoplanBandClass = Choose(clusterBandClass, "Upper 700 MHz", "Cellular")
    
    
    
    Dim rngLocationTable As Range
    Dim locationCount As Long
    
    Dim locationID As String
    Dim locationLat As Double, locationLon As Double
    Dim locationRow As Long
    
    Set rngLocationTable = ThisWorkbook.Names("MULTI_LOC_LOCATION_TABLE").RefersToRange
    
    
    locationCount = Range_LastRow(rngLocationTable)
    
    
    Dim nIdx As Long
    
    Dim cellList() As CellSite
    Dim neighbors() As NeighborCell
    
    
    Dim allNeighbors() As CellSite
    Dim allNeighbors_Count As Long
    Dim allNeighbors_offloadSectorsByLocationIndex() As Byte
    
    Dim allNeighbors_Old_Count As Long
    Dim allNeighborsNextFreeIdx As Long, allNeighborsIdx As Long
    
    Const ALPHA_BIT = 1
    Const BETA_BIT = 2
    Const GAMMA_BIT = 4
    
    Dim isNeighborInAllNeighborsList As Boolean
    Dim neighborNotInAllNeighborsListCount As Long
    
    
    cellList = LoadCellListFromGeoPlanReport(geoPlanDataFile, geoplanBandClass)
    
    Dim cellCount As Long
    
    On Error Resume Next
    cellCount = 0
    cellCount = UBound(cellList) - LBound(cellList) + 1
    On Error GoTo 0
    
    If cellCount = 0 Then
        MsgBox "Could not load cells from file: " & geoPlanDataFile
        Exit Sub
    End If
    
    For locationRow = 1 To locationCount
        locationID = rngLocationTable.Cells(locationRow, 1)
        locationLat = rngLocationTable.Cells(locationRow, 2)
        locationLon = rngLocationTable.Cells(locationRow, 3)
        
        
        
        neighbors = Calc1stTierNeighbors(cellList, locationLat, locationLon, maxNeighborDistance, True)
        
        
        Debug.Print locationID & " (" & locationLat & ", " & locationLon & ")"
        
        
        neighborNotInAllNeighborsListCount = 0
        
        ' ------------------------------------------------------------------------------------
        ' Begin - Add each neighbor to allNeighborsList, if they do not exist there already. Record all offload sectors
        ' ------------------------------------------------------------------------------------
        
        
        Dim neighborCount As Long
        
        On Error Resume Next
        neighborCount = 0
        neighborCount = UBound(neighbors) - LBound(neighbors) + 1
        On Error GoTo 0
        
        If neighborCount = 0 Then
            MsgBox "No 1st tier neighbors found for cluster"
            Exit Sub
        End If
        
        
        ' Start by counting the number of neighbors in list
        For nIdx = LBound(neighbors) To UBound(neighbors)
            isNeighborInAllNeighborsList = False
            
            For i = 1 To allNeighbors_Count
                If allNeighbors(i).UID = neighbors(nIdx).cell.UID Then
                    isNeighborInAllNeighborsList = True
                    Exit For
                End If
            Next i
            
            If isNeighborInAllNeighborsList = False Then neighborNotInAllNeighborsListCount = neighborNotInAllNeighborsListCount + 1
        Next nIdx
        
        ' Start adding from the previous count
        allNeighbors_Old_Count = allNeighbors_Count
        allNeighborsNextFreeIdx = allNeighbors_Count + 1 ' Next index where we can add cell
        
        If neighborNotInAllNeighborsListCount > 0 Then
            allNeighbors_Count = allNeighbors_Old_Count + neighborNotInAllNeighborsListCount
            
            ReDim Preserve allNeighbors(1 To allNeighbors_Count)
            ReDim Preserve allNeighbors_offloadSectorsByLocationIndex(1 To locationCount, 1 To allNeighbors_Count) As Byte
        End If
        
        ' Now add to all neighbors list and record offload sectors for each location
        For nIdx = LBound(neighbors) To UBound(neighbors)
            
            ' Does neighbor already exist in all neighbors list? If so, return index
            allNeighborsIdx = 0
            
            For i = 1 To allNeighbors_Old_Count 'NOTE: we use the old count here
                If allNeighbors(i).UID = neighbors(nIdx).cell.UID Then
                    allNeighborsIdx = i
                    Exit For
                End If
            Next i
            
            ' Debug
            'Dim isExistingNeighbor As Boolean, offSectorStr As String
            'isExistingNeighbor = allNeighborsIdx > 0
            'offSectorStr = ""
            'If neighbors(nIdx).is_offload_alpha Then offSectorStr = offSectorStr & "Alpha, "
            'If neighbors(nIdx).is_offload_beta Then offSectorStr = offSectorStr & "Beta, "
            'If neighbors(nIdx).is_offload_gamma Then offSectorStr = offSectorStr & "Gamma, "
            
            'Debug.Print " - " & IIf(isExistingNeighbor, "EXIST: ", "NOT_EXIST: ") & neighbors(nIdx).cell.Name & " (" & neighbors(nIdx).cell.UID & ") " & offSectorStr
            'Debug
            
            ' neighbor is not found in all neighbors list (no index found) -> add neighbor to list
            If allNeighborsIdx = 0 Then
                allNeighbors(allNeighborsNextFreeIdx) = neighbors(nIdx).cell
                allNeighbors_offloadSectorsByLocationIndex(locationRow, allNeighborsNextFreeIdx) = 0
                
                allNeighborsIdx = allNeighborsNextFreeIdx
                allNeighborsNextFreeIdx = allNeighborsNextFreeIdx + 1
            End If
            
            
            If neighbors(nIdx).is_offload_alpha Then allNeighbors_offloadSectorsByLocationIndex(locationRow, allNeighborsIdx) = allNeighbors_offloadSectorsByLocationIndex(locationRow, allNeighborsIdx) + ALPHA_BIT
            If neighbors(nIdx).is_offload_beta Then allNeighbors_offloadSectorsByLocationIndex(locationRow, allNeighborsIdx) = allNeighbors_offloadSectorsByLocationIndex(locationRow, allNeighborsIdx) + BETA_BIT
            If neighbors(nIdx).is_offload_gamma Then allNeighbors_offloadSectorsByLocationIndex(locationRow, allNeighborsIdx) = allNeighbors_offloadSectorsByLocationIndex(locationRow, allNeighborsIdx) + GAMMA_BIT
  
        Next nIdx
        ' ------------------------------------------------------------------------------------
        ' End - Add each neighbor to allNeighborsList, if they do not exist there already
        ' ------------------------------------------------------------------------------------
        
    Next locationRow
    
    
    ' Begin - Sort neghbors using their indices (NORTH -> SOUTH)
    Dim allNeighborsSortedIdx() As Long, idx As Long
    ReDim allNeighborsSortedIdx(1 To allNeighbors_Count) As Long
    
    For i = 1 To allNeighbors_Count: allNeighborsSortedIdx(i) = i: Next i

    'For i = 1 To allNeighbors_Count
    '    For j = i + 1 To allNeighbors_Count
    '        If allNeighbors(allNeighborsSortedIdx(j)).lat > allNeighbors(allNeighborsSortedIdx(i)).lat Then
    '            idx = allNeighborsSortedIdx(i)
    '            allNeighborsSortedIdx(i) = allNeighborsSortedIdx(j)
    '            allNeighborsSortedIdx(j) = idx
    '        End If
    '    Next j
    'Next i

    
    Dim outRng As Range
    
    ' Start output immediately to the right and 2 rows up from location table (two rows for header)
    Set outRng = rngLocationTable.Offset(-2, rngLocationTable.columns.count).Resize(locationCount + 2, 3 * 40)
    
    outRng.ClearContents
    
    
    Excel_AppUpdates_Disable
    
    For i = 1 To allNeighbors_Count
        idx = allNeighborsSortedIdx(i)
        
        outRng.Cells(1, (i - 1) * 3 + 1) = allNeighbors(idx).Name
        outRng.Cells(2, (i - 1) * 3 + 1) = allNeighbors(idx).UID & "-1"
        outRng.Cells(2, (i - 1) * 3 + 2) = allNeighbors(idx).UID & "-2"
        outRng.Cells(2, (i - 1) * 3 + 3) = allNeighbors(idx).UID & "-3"
        
    Next i
    
    Set outRng = outRng.Offset(2)
    
    
    For i = 1 To allNeighbors_Count
        idx = allNeighborsSortedIdx(i)
        
        For j = 1 To locationCount
            If allNeighbors_offloadSectorsByLocationIndex(j, idx) And ALPHA_BIT Then outRng.Cells(j, (i - 1) * 3 + 1) = "X"
            If allNeighbors_offloadSectorsByLocationIndex(j, idx) And BETA_BIT Then outRng.Cells(j, (i - 1) * 3 + 2) = "X"
            If allNeighbors_offloadSectorsByLocationIndex(j, idx) And GAMMA_BIT Then outRng.Cells(j, (i - 1) * 3 + 3) = "X"
        Next j
    Next i

    Excel_AppUpdates_Restore



    

End Sub

