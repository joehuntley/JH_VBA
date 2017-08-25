Attribute VB_Name = "JH_ReportingAutomation"
Option Explicit


' JH_ReportingAutomation
' ------------------------------------------------------------------------------------------------------------------
' VBA functions to interact with and process data from VZW reporting tools such as ALPT, NCWS, MPT
'
' Joseph Huntley (joseph.huntley@vzw.com)
' ------------------------------------------------------------------------------------------------------------------
' Changelog:
'
' 2016-10-02 (joe h)    - Integrated MTSO_Reference_Table_XXX functions
'                       - Updated Cluster_CellGroupReport_1XDO to work with clusters defined with SysID
'                       - Updated ClusterDef_Expand to work with clusters defined with SysID
' 2015-06-22 (joe h)    - Added ALPT_CustomReport_Cluster, ALPT_CustomReport_EnodeList
' 2015-05-13 (joe h)    - Event_FilterWorksheetByDateRange: Allow specifying columns by name
'                       - ClusterDef_LTE_Add_AWS_eNodeBs: Fixed sector counting bug
' 2015-05-09 (joe h)    - Added Event_Download_LTE_Data()
' 2015-05-01 (joe h)    - Cluster_CellGroupReport_1XDO: Added NCWS_1X_CELL query.
'                       - Added event related functions: Event_CreateDateRanges, Event_FilterWorksheetByDateRange, Event_ParseDaysOfWeek
'                       - Fixed ClusterDef_LTE_Add_AWS_eNodeBs to allow up to 6 sectors per enode B
' 2015-04-02 (joe h)    - Cluster_CellGroupReport_1XDO: Fixed param check
' 2015 (joe h)          - Initial version
' ------------------------------------------------------------------------------------------------------------------

' -----------------------------------------------------------------------------------------------
' Begin - ALPT Constants/Types/Enums
' -----------------------------------------------------------------------------------------------
Public Type ALPT_Session
    user As String
    pass As String
    
    BASE_URI As String
    
    loggedIn As Boolean
    
    httpReq As WinHttp.WinHttpRequest
End Type

Private Type ALPT_ElementListItem
    value As String
    description As String
End Type

' ALPT Structure Ad-Hoc Query Result
Public Type ALPT_SAH_Result
    dataResourceName As String
    reportID As Long
End Type


' ALPT Structured Ad-Hoc Report Level:
' - Found in HTML source of http://alpt.vh.eng.vzwcorp.com:8282/alte/sah.htm
Enum ALPT_SAH_Report_Level
    MME_Pool = 1
    mme = 3
    MME_Service = 9
    MME_Service_Member = 10
    MME_Cabinet = 14
    MME_Shelf = 15
    MME_Card = 16
    MME_Host = 17
    MME_Core = 18
    MME_Svc_Type = 19
    SGW_ = 2
    SGW_Slot = 11
    SGW_MDA = 12
    SGW_VRID = 13
    SGW_KCI = 20
    SGW_KPI = 21
    SGW_GroupName = 22
    SGW_Group = 23
    SGW_Address = 24
    SGW_Port = 25
    PGW = 26
    PGW_APN = 29
    PGW_Slot = 27
    PGW_Port = 28
    Region = 4
    market = 5
    Cell_Group = 6
    enodeB = 7
    Carrier_Rollup = 30
    euTrancell = 8
    Sector_Carrier = 33
    Market_Carrier = 31
    Region_Carrier = 32
    IntraFrqNcellRel = 34
    AntennaPort = 35
    LteNeighboringFreqConf = 36
    Roll_Up = 37
End Enum


' ALPT Structured Ad-Hoc Report Type:
' - Found in HTML source of http://alpt.vh.eng.vzwcorp.com:8282/alte/sah.htm
Enum ALPT_SAH_Report_Type
    Fifteen_Min_Totals = 1
    Hourly_Agg_Totals = 34
    Hourly_Totals = 2
    Daily_Totals = 3
    Agg_Totals = 16
    Fifteen_Min_Agg_Totals = 35
    MME_POOL_DATAVOL_BH = 25
    MME_POOL_DLUL_RLC_BH = 20
    MME_POOL_DL_RLC_BH = 22
    MME_POOL_InitErabSetupRqst_BH = 7
    MME_POOL_RrcConn_BH = 8
    MMEPool_QCI1_Atts_BH = 40
    MMEPool_VoLTE_MOU_BH = 46
    Market_InitErabSetupRqst_BH = 15
    Market_RrcConn_BH = 6
    eNodeB_InitErabSetupRqst_BH = 12
    eNodeB_RrcConn_BH = 5
    eNodeB_QCI1_Atts_BH = 38
    eNodeB_VoLTE_MOU_BH = 44
    euCell_DL_RLC_Layer_MByte_BH = 17
    euCell_InitErabSetupRqst_BH = 13
    euCell_RrcConn_BH = 4
    euCell_VoLTE_MOU_BH = 42
    euCell_QCI1_Atts_BH = 36
    Agg_MME_POOL_DATAVOL_BH = 33
    Agg_MME_POOL_DLUL_RLC_BH = 32
    Agg_MME_POOL_DL_RLC_BH = 23
    Agg_MME_POOL_InitErabSetupRqst_BH = 29
    Agg_MME_POOL_RrcConn_BH = 19
    Agg_MMEPool_QCI1_Atts_BH = 41
    Agg_MMEPool_VoLTE_MOU_BH = 47
    Agg_Market_InitErabSetupRqst_BH = 31
    Agg_Market_RrcConn_BH = 28
    Agg_eNodeB_InitErabSetupRqst = 26
    Agg_eNodeB_RrcConn_BH = 18
    Agg_eNodeB_QCI1_Atts_BH = 39
    Agg_eNodeB_VoLTE_MOU_BH = 45
    Agg_euCell_DL_RLC_Layer_MByte_BH = 24
    Agg_euCell_InitErabSetupRqst_BH = 30
    Agg_euCell_RrcConn_BH = 27
    Agg_euCell_QCI1_Atts_BH = 37
    Agg_euCell_VoLTE_MOU_BH = 43
End Enum

' ALPT Structured Ad-Hoc Report Content:
' - Found in HTML source of http://alpt.vh.eng.vzwcorp.com:8282/alte/sah.htm
Public Enum ALPT_SAH_Report_Content
    HQ_LTE_KPIs = 31
    HQ_eNodeB_KPIs = 29
    eNodeB_All_KPIs_Perf = 1
    eNodeB_Mobility_Perf = 2
    eNodeB_Retainability_Perf = 3
    eNodeB_Accessibility_Perf = 4
    eNodeB_Throughput_Perf = 5
    eNodeB_Config_AntPort = 65
    eNodeB_Config_CAC = 66
    eNodeB_Config_Equip = 67
    eNodeB_Config_VoIP = 68
    eNodeB_RF_Quality_Perf = 6
    eNodeB_Geographic_Info = 42
    eNodeB_Irisview_VoLTE_QOS = 54
    eNodeB_Irisview_VoLTE_SIP = 55
    eNodeB_IF_HO = 56
    HQ_MME_KPIs = 30
    MME_Mobility = 7
    MME_Accessibility = 8
    MME_Alarms = 9
    MME_Authentication = 10
    MME_CPU_Usage = 11
    MME_CPU_Usage_by_Svc_Type = 12
    MME_Disk_IO_Rate = 13
    MME_File_Sys_Usage = 14
    MME_Global_Disk_ = 15
    MME_Mem_Usage_by_Svc_Type = 16
    MME_UE_Connections = 17
    MME_Mem_Usage = 18
    MME_SG_Report = 19
    VoLTE_MME_Rpt = 50
    SGW_BHCA_KPIs = 47
    SGW_CPU_Utilization = 20
    SGW_CPU_Utilization_MDA = 40
    SGW_CPU_Utilization_CPISA = 45
    SGW_CPU_Utilization_CPM = 46
    SGW_Capacity_Bearer = 21
    SGW_Capacity_Control_Plane = 22
    SGW_IP_Utilization = 23
    SGW_Slot_Level_stats = 24
    SGW_Mobility_25 = 25
    SGW_Paging = 26
    SGW_Bearer_Report = 27
    SGW_Control_Plane_Report = 28
    SGW_KPIs = 70
    SGW_Accessibility = 71
    SGW_Accessibility_mda = 72
    SGW_Accessibility_port = 73
    SGW_Mobility_74 = 74
    SGW_Capacity = 75
    SGW_CPU_Utilization_CP_ISA = 76
    Cisco_SGW_KPI_Report = 32
    eNodebAllPegs = 33
    Cisco_PGW_Bearers_Report = 39
    Cisco_PGW_Bearers_by_APN_Report = 51
    Cisco_PGW_per_QCI_by_APN_Report = 52
    Cisco_PGW_Tput_KPIs_by_APN_Report = 53
    Cisco_PGW_CPU_Memory_by_Slot = 34
    Cisco_PGW_QCI_Report = 35
    Cisco_PGW_Sessions_Report = 36
    Cisco_PGW_Handover_Report = 37
    Cisco_PGW_Port_Level_Stats = 38
    eNodeB_UL_Noise = 44
    eNodeB_CQI_Perf_Rev13_3 = 69
    eNodeB_CQI_Perf = 43
    VoLTE_eNB_Rpt = 49
    Eutrancell_RSSI = 48
    IntraFrqNcellRel_Content = 58
    Eutrancell_TxPwr = 57
    Eutrancell_Config_Freq_AntPort = 59
    Eutrancell_Config_L2 = 60
    Eutrancell_Config_Cell = 61
    Eutrancell_Config_Neighbor = 62
    Eutrancell_Config_CAC = 63
    Eutrancell_Config_SysInfo = 64
End Enum

' ALPT Internal Region IDs:
' - Found in: http://alpt.vh.eng.vzwcorp.com:8282/alte/rptWebServiceHelp.htm
Private Enum ALPT_Region_ID
    Southwest = 3
    New_York_Metro = 8
    Washington_Baltimore_Virginia = 10
    Mountain = 2
    Mountain_LRA = 68
    South_Central = 16
    South_LRA = 28
    Ohio_Pennsylvania = 21
    Philadelphia_TriState = 9
    Houston_Gulf_Coast = 11
    Central_Texas = 12
    New_England = 6
    Upstate_New_York = 7
End Enum


' ALPT Internal Market IDs:
' - Found in HTML source of http://alpt.vh.eng.vzwcorp.com:8282/alte/myPrefs.htm
' - Note: These are the market identifiers used internally by ALPT. They are not the same as the LTE Market ID
Private Enum ALPT_Internal_Market_ID
    ' Central TX
    LA_Shreveport = 103
    TX_Austin = 106
    TX_Dallas = 102
    TX_East_TX_Shreveport = 104
    TX_Fort_Worth = 101
    TX_Lubbock = 107
    TX_Midland = 108
    TX_North_Texas = 288
    TX_Rio_Grande_Valley = 290
    TX_San_Antonio_Schertz = 105
    ' Houston/Gulf Coast
    FL_Pensacola = 99
    LA_Baton_Rouge = 97
    LA_Covington = 98
    MS_Jackson = 100
    MS_Jackson_Covington = 96
    TX_Cicero = 92
    TX_Copperfield = 93
    TX_Lake_Charles_West_Park = 95
    TX_WestPark = 94
    ' Mountain
    CO_Denver_Aurora1 = 15
    CO_Denver_Aurora2 = 14
    CO_Denver_Clinton = 12
    CO_Denver_Westminster = 13
    ' Mountain LRA
    ID_Idaho_Boise = 9
    MT_Montana_Helena = 8
    UT_Salt_Lake_Kearns = 10
    UT_Salt_Lake_W_Jordan = 11
    ID_LRA_Custer = 254
    MT_LRA_MidRiver = 328
    UT_LRA_Strata = 252
    ' New England
    CT_Central_Counties_Wallingford_1 = 49
    CT_Fairfield_County_Wallingford_2 = 50
    CT_Hartford_County_Windsor_2 = 52
    MA_Boston_W_Roxbury_1 = 42
    MA_Central_MA_Westboro = 45
    MA_Northeastern_MA_Billerica = 44
    MA_Southeastern_MA_W_Roxbury_2 = 43
    ME_Maine_Hooksett = 48
    NH_New_Hampshire_Hooksett = 47
    RI_and_Bristol_County_MA_Taunton = 46
    VT_and_Western_MA_Counties_Windsor_1 = 51
    ' New York Metro
    NNJ_Central_Wayne_1 = 71
    NNJ_East_Jersey_City_2 = 67
    NNJ_Middlesex_Branchburg_3 = 72
    NNJ_NorthEast_Jersey_City_1 = 66
    NNJ_NorthWest_Wayne_2 = 70
    NNJ_South_Branchburg_2 = 69
    NNJ_West_Branchburg_1 = 68
    NYM_Bronx_Upper_Manhattan_WNyack2 = 59
    NYM_Lower_Manhattan_Mineola2 = 64
    NYM_Midtown_Manhattan_Yonkers = 60
    NYM_Nassau_Farmingdale2 = 63
    NYM_Queens_Whitestone_2 = 62
    NYM_Staten_Island_Brooklyn_Whitestone_1 = 61
    NYM_Suffolk_Farmingdale_1 = 65
    NYM_Westch_Rockld_Putnam_WNyack1 = 58
    ' Ohio/Pennsylvania
    OH_Akron_1 = 190
    OH_Akron_2 = 191
    OH_Cincinnati = 193
    OH_Cincinnati_Duff_1 = 194
    OH_Cleveland_Cleveland_1 = 188
    OH_Columbus = 192
    OH_Columbus_Lewis_Center_1 = 196
    OH_Columbus_Lewis_Center_2 = 197
    OH_Dayton_Duff_2 = 195
    OH_NW_Ohio_Maumee = 198
    OH_Toledo_Cleveland_2 = 189
    PA_Pittsburgh_Bridgeville = 186
    PA_Pittsburgh_Bridgeville_2 = 187
    PA_Pittsburgh_Johnstown = 185
    PA_Pittsburgh_Pittsburgh = 184
    WV_Huntington_St_Clairesville = 199
    ' Philadelphia Tri-State
    DE_Wilmington = 79
    NJ_Burlington_Maple_Shade = 77
    NJ_Vineland_Wilmington_2 = 76
    PA_Harrisburg_Harrisburg = 74
    PA_Lehigh_Valley_Plymouth_Meeting_2 = 80
    PA_Philadelphia_Philadelphia = 75
    PA_Plymouth_Meeting = 78
    PA_Scranton_WilkesBarre_Pittston = 73
    ' South Central
    AR_Greater_Arkansas = 11000
    AR_LittleRock = 139
    AR_Northwest_AR_Fort_Smith = 140
    OK_Tulsa = 141
    TN_Memphis_WestTenn = 138
    ' South LRA
    OK_LRA_Cross = 250
    OK_LRA_Pioneer = 248
    OK_LRA_Pioneer_2 = 1010
    ' Southwest
    AZ_Phoenix_Gilbert = 17
    AZ_Phoenix_PHX = 18
    AZ_Phoenix_Tempe = 20
    AZ_Tucson = 19
    NM_Albuquerque = 21
    NV_Las_Vegas = 16
    TX_El_Paso = 22
    ' Upstate New York
    NY_Albany_North_Greenbush_1 = 54
    NY_Buffalo = 53
    NY_North_Greenbush_2 = 57
    NY_Rochester = 56
    NY_Syracuse = 55
    ' Washington/Baltimore/Virginia
    DC_Washington_Adelphi = 82
    MD_Baltimore_Catonsville = 83
    MD_Frederick_Woodlawn_2 = 84
    MD_Salisbury = 85
    MD_Silver_Spring_Annpolis_Junction = 81
    MD_Southern_Maryland_Woodlawn_1 = 86
    VA_Chantilly_2 = 208
    VA_Charlottesville_Goods_Bridge = 90
    VA_Fairfax_Chantilly = 87
    VA_Norfolk_Lee_Hall = 88
    VA_Richmond = 89
    VA_Roanoke = 91
End Enum


Private Type ALPT_LookupTable_MarketItem
    ALPT_RegionID As ALPT_Region_ID
    ALPT_InternalMarketID As ALPT_Internal_Market_ID
    LTE_Market_ID As Integer
End Type


Private LookupTable_ALPT_Markets_Initialized As Boolean
Private LookupTable_ALPT_Markets() As ALPT_LookupTable_MarketItem
' -----------------------------------------------------------------------------------------------
' End - ALPT Constants/Types/Enums
' -----------------------------------------------------------------------------------------------


Type MTSO_Reference_Table_Item
    LTE_Market_ID As Integer
    SysID As Integer
    SN As Integer
    ECP As Integer
    
    ReportingTool_1XDO As String ' MPT or NCWS
    
    NCWS_1XDO_CellGroupName As String
    MPT_1XDO_MarketName As String
End Type



' Function ALPT_eNodeList_To_CellIDs
' --------------------------------------------
' ALPT References cells/eNBs and LTE markets by an internal ID, which is separate from the enodeB ID and
' LTE market ID, respectively. The eNB ID is provided as part of dropdown lists on the webpage. This function
' receives a list of enodeB IDs and performs the following actions:
'
' 1) Create a unique list of LTE Market IDs from each eNodeB ID.
' 2) Map each LTE Market ID their equivalent ALPT Internal Market ID
' 3) Query ALPT servers for all eNodeBs within the markets (cell ID + enodeid_description)
' 4) Map the provided enodeB IDs to their equivalent ALPT internal Cell IDs
'
' Note: If a provided eNodeB ID cannot be found in the ALPT list, then the returned ALPT Cell ID is 0
Public Function ALPT_eNodeList_To_CellIDs(session As ALPT_Session, enodeList() As Long) As Long()

    'Dim session As ALPT_Session :session = ALPT_NewSession()
    'Dim clusterDefList As Variant: clusterDefList = "098048-2;098317-2;098016-2;098600-3"
    'Dim enodeList() As Long: enodeList = ClusterDef_Extract_eNodeList(clusterDefList)
    
    ' check if enodelist is empty
    If Array_Count(enodeList) = 0 Then
        ALPT_eNodeList_To_CellIDs = enodeList()
        Exit Function
    End If
        
    
    ' -----------------------------------------------------------------------------------------
    ' Begin - Use LTE market IDs from the enodeB IDs to create list of ALPT internal market IDs
    ' -----------------------------------------------------------------------------------------
    Dim LTE_Mkt_ID_alreadyProcessedArr() As Long
    Dim LTE_Mkt_ID_alreadyProcessedCount As Long
    Dim LTE_Mkt_ID_alreadyProcessed As Boolean
    
    Dim foundAlptMarket As Boolean
    Dim alptInternalMarketIDs As String 'market IDs separated by carets (^) eg. 66^68^70
    Dim enodeID As Long, LTE_Mkt_ID As Long

    Dim i As Integer, j As Integer, k As Integer
    
    
    LookupTable_ALPT_Markets_Build

   
    
    alptInternalMarketIDs = ""
    LTE_Mkt_ID_alreadyProcessedCount = 0
    
    For i = LBound(enodeList) To UBound(enodeList)
        enodeID = enodeList(i)
        
        LTE_Mkt_ID = Int(enodeID / 1000) ' Convert 999123 to 999
        'LTE_Mkt_ID = IIf(LTE_Mkt_ID >= 300, LTE_Mkt_ID - 300, LTE_Mkt_ID) ' Convert AWS Market IDs to true LTE market IDs
        LTE_Mkt_ID = LTE_Mkt_ID Mod 300 ' Convert AWS Market IDs to true LTE market IDs
        
        If LTE_Mkt_ID = 0 Then Err.Raise -1, , "Invalid LTE Market ID for eNodeB: " & enodeID
        
        If LTE_Mkt_ID > 0 Then
            ' Check if we already processed an eNodeB with the same market ID
            LTE_Mkt_ID_alreadyProcessed = False
            
            For j = 1 To LTE_Mkt_ID_alreadyProcessedCount
                If LTE_Mkt_ID_alreadyProcessedArr(j) = LTE_Mkt_ID Then
                    LTE_Mkt_ID_alreadyProcessed = True
                    Exit For
                  End If
            Next j
            
            
            
            If LTE_Mkt_ID_alreadyProcessed = False Then
                ' Add current LTE market ID to array to indicate the ALPT internal market ID has already been looked up
                LTE_Mkt_ID_alreadyProcessedCount = LTE_Mkt_ID_alreadyProcessedCount + 1
                ReDim Preserve LTE_Mkt_ID_alreadyProcessedArr(1 To LTE_Mkt_ID_alreadyProcessedCount) As Long
                LTE_Mkt_ID_alreadyProcessedArr(LTE_Mkt_ID_alreadyProcessedCount) = LTE_Mkt_ID
                
                
                ' Find corresponding internal ALPT Market ID corresponding to the LTE_Market_ID
                foundAlptMarket = False
                
                For j = LBound(LookupTable_ALPT_Markets) To UBound(LookupTable_ALPT_Markets)
                    If LookupTable_ALPT_Markets(j).LTE_Market_ID = LTE_Mkt_ID Then
                        alptInternalMarketIDs = alptInternalMarketIDs & LookupTable_ALPT_Markets(j).ALPT_InternalMarketID & "^"
        
                        foundAlptMarket = True
                        Exit For
                    End If
                Next j
                
                If foundAlptMarket = False Then Err.Raise -1, , "Cannot map LTE Market ID (" & LTE_Mkt_ID & ") to ALPT Internal Market ID for eNodeB: " & enodeID
           
            End If
            
        End If
        
    Next i
    
    
    ' remove trailing caret (^), if it exists. alptInternalMarketIDs will contain all the unique markets separated by a caret (^)
    If Len(alptInternalMarketIDs) > 0 Then alptInternalMarketIDs = Left$(alptInternalMarketIDs, Len(alptInternalMarketIDs) - 1)
    ' -----------------------------------------------------------------------------------------
    ' End - Use LTE market IDs from the enodeB IDs to create list of ALPT internal market IDs
    ' -----------------------------------------------------------------------------------------


    ' -----------------------------------------------------------------------------------------
    ' Begin - Retrive enodeB list from ALPT. Map our enodeB list to ALPT Internal Cell IDs
    ' -----------------------------------------------------------------------------------------
    ALPT_Session_Login session
   
   
   
    Dim queryString As String
    Dim lptEnodeList() As ALPT_ElementListItem
    Dim lptEnodeListCount As Long
    
    
    queryString = URL_BuildQueryString("userName", alptInternalMarketIDs, "type", "enodeb", "subType", "markets", "sortBy", "nbr", "user", session.user)


    lptEnodeListCount = ALPT_GetElementList(session, queryString, lptEnodeList)
    
    
    
    Dim lptEnodeId As Long
    
    Dim lptCellList() As Long
    Dim lptCellListIdx As Long
         
    
    
    ReDim lptCellList(LBound(enodeList) To UBound(enodeList)) As Long
    
    
    ' Set all items in cell ID list to zero (not found)
    For j = LBound(enodeList) To UBound(enodeList): lptCellList(j) = 0: Next j
    
    ' Loop through ALPT enode list elements and set cell ID
    For i = 1 To lptEnodeListCount
        ' Description is in the format of 099123_SITE_NAME
        ' Value is the Cell ID
        lptEnodeId = CLng(Left(lptEnodeList(i).description, InStr(1, lptEnodeList(i).description, "_") - 1))
            
        For j = LBound(enodeList) To UBound(enodeList)
            If lptCellList(j) = 0 And enodeList(j) = lptEnodeId Then
                lptCellList(j) = CLng(lptEnodeList(i).value)
            End If
        Next j
    Next i
    ' -----------------------------------------------------------------------------------------
    ' End - Retrive enodeB list from ALPT. Map our enodeB list to ALPT Internal Cell IDs
    ' -----------------------------------------------------------------------------------------

   
   ALPT_eNodeList_To_CellIDs = lptCellList
    
End Function


Private Sub LookupTable_ALPT_Markets_Build()

    If LookupTable_ALPT_Markets_Initialized Then Exit Sub
    
    
    LookupTable_ALPT_Markets_Initialized = True
    
    
    Const MARKET_COUNT As Integer = 112 ' Total # of markets per ALPT
            
    ReDim LookupTable_ALPT_Markets(1 To MARKET_COUNT) As ALPT_LookupTable_MarketItem
    
    
    Dim i As Integer
    
    
    i = 0
        
    i = i + 1:  LookupTable_ALPT_Markets_PopulateItem i, ALPT_Region_ID.Central_Texas, ALPT_Internal_Market_ID.LA_Shreveport, 134
    i = i + 1:  LookupTable_ALPT_Markets_PopulateItem i, ALPT_Region_ID.Central_Texas, ALPT_Internal_Market_ID.TX_Austin, 137
    i = i + 1:  LookupTable_ALPT_Markets_PopulateItem i, ALPT_Region_ID.Central_Texas, ALPT_Internal_Market_ID.TX_Dallas, 133
    i = i + 1:  LookupTable_ALPT_Markets_PopulateItem i, ALPT_Region_ID.Central_Texas, ALPT_Internal_Market_ID.TX_East_TX_Shreveport, 135
    i = i + 1:  LookupTable_ALPT_Markets_PopulateItem i, ALPT_Region_ID.Central_Texas, ALPT_Internal_Market_ID.TX_Fort_Worth, 132
    i = i + 1:  LookupTable_ALPT_Markets_PopulateItem i, ALPT_Region_ID.Central_Texas, ALPT_Internal_Market_ID.TX_Lubbock, 138
    i = i + 1:  LookupTable_ALPT_Markets_PopulateItem i, ALPT_Region_ID.Central_Texas, ALPT_Internal_Market_ID.TX_Midland, 139
    i = i + 1:  LookupTable_ALPT_Markets_PopulateItem i, ALPT_Region_ID.Central_Texas, ALPT_Internal_Market_ID.TX_North_Texas, 131
    i = i + 1:  LookupTable_ALPT_Markets_PopulateItem i, ALPT_Region_ID.Central_Texas, ALPT_Internal_Market_ID.TX_Rio_Grande_Valley, 140
    i = i + 1:  LookupTable_ALPT_Markets_PopulateItem i, ALPT_Region_ID.Central_Texas, ALPT_Internal_Market_ID.TX_San_Antonio_Schertz, 136
    i = i + 1:  LookupTable_ALPT_Markets_PopulateItem i, ALPT_Region_ID.Houston_Gulf_Coast, ALPT_Internal_Market_ID.FL_Pensacola, 127
    i = i + 1:  LookupTable_ALPT_Markets_PopulateItem i, ALPT_Region_ID.Houston_Gulf_Coast, ALPT_Internal_Market_ID.LA_Baton_Rouge, 125
    i = i + 1:  LookupTable_ALPT_Markets_PopulateItem i, ALPT_Region_ID.Houston_Gulf_Coast, ALPT_Internal_Market_ID.LA_Covington, 126
    i = i + 1:  LookupTable_ALPT_Markets_PopulateItem i, ALPT_Region_ID.Houston_Gulf_Coast, ALPT_Internal_Market_ID.MS_Jackson, 128
    i = i + 1:  LookupTable_ALPT_Markets_PopulateItem i, ALPT_Region_ID.Houston_Gulf_Coast, ALPT_Internal_Market_ID.MS_Jackson_Covington, 124
    i = i + 1:  LookupTable_ALPT_Markets_PopulateItem i, ALPT_Region_ID.Houston_Gulf_Coast, ALPT_Internal_Market_ID.TX_Cicero, 120
    i = i + 1:  LookupTable_ALPT_Markets_PopulateItem i, ALPT_Region_ID.Houston_Gulf_Coast, ALPT_Internal_Market_ID.TX_Copperfield, 121
    i = i + 1:  LookupTable_ALPT_Markets_PopulateItem i, ALPT_Region_ID.Houston_Gulf_Coast, ALPT_Internal_Market_ID.TX_Lake_Charles_West_Park, 123
    i = i + 1:  LookupTable_ALPT_Markets_PopulateItem i, ALPT_Region_ID.Houston_Gulf_Coast, ALPT_Internal_Market_ID.TX_WestPark, 122
    i = i + 1:  LookupTable_ALPT_Markets_PopulateItem i, ALPT_Region_ID.Mountain, ALPT_Internal_Market_ID.CO_Denver_Aurora1, 17
    i = i + 1:  LookupTable_ALPT_Markets_PopulateItem i, ALPT_Region_ID.Mountain, ALPT_Internal_Market_ID.CO_Denver_Aurora2, 16
    i = i + 1:  LookupTable_ALPT_Markets_PopulateItem i, ALPT_Region_ID.Mountain, ALPT_Internal_Market_ID.CO_Denver_Clinton, 14
    i = i + 1:  LookupTable_ALPT_Markets_PopulateItem i, ALPT_Region_ID.Mountain, ALPT_Internal_Market_ID.CO_Denver_Westminster, 15
    i = i + 1:  LookupTable_ALPT_Markets_PopulateItem i, ALPT_Region_ID.Mountain, ALPT_Internal_Market_ID.ID_Idaho_Boise, 11
    i = i + 1:  LookupTable_ALPT_Markets_PopulateItem i, ALPT_Region_ID.Mountain, ALPT_Internal_Market_ID.MT_Montana_Helena, 10
    i = i + 1:  LookupTable_ALPT_Markets_PopulateItem i, ALPT_Region_ID.Mountain, ALPT_Internal_Market_ID.UT_Salt_Lake_Kearns, 12
    i = i + 1:  LookupTable_ALPT_Markets_PopulateItem i, ALPT_Region_ID.Mountain, ALPT_Internal_Market_ID.UT_Salt_Lake_W_Jordan, 13
    i = i + 1:  LookupTable_ALPT_Markets_PopulateItem i, ALPT_Region_ID.Mountain_LRA, ALPT_Internal_Market_ID.ID_LRA_Custer, 9610
    i = i + 1:  LookupTable_ALPT_Markets_PopulateItem i, ALPT_Region_ID.Mountain_LRA, ALPT_Internal_Market_ID.MT_LRA_MidRiver, 9620
    i = i + 1:  LookupTable_ALPT_Markets_PopulateItem i, ALPT_Region_ID.Mountain_LRA, ALPT_Internal_Market_ID.UT_LRA_Strata, 9600
    i = i + 1:  LookupTable_ALPT_Markets_PopulateItem i, ALPT_Region_ID.New_England, ALPT_Internal_Market_ID.CT_Central_Counties_Wallingford_1, 64
    i = i + 1:  LookupTable_ALPT_Markets_PopulateItem i, ALPT_Region_ID.New_England, ALPT_Internal_Market_ID.CT_Fairfield_County_Wallingford_2, 65
    i = i + 1:  LookupTable_ALPT_Markets_PopulateItem i, ALPT_Region_ID.New_England, ALPT_Internal_Market_ID.CT_Hartford_County_Windsor_2, 67
    i = i + 1:  LookupTable_ALPT_Markets_PopulateItem i, ALPT_Region_ID.New_England, ALPT_Internal_Market_ID.MA_Boston_W_Roxbury_1, 56
    i = i + 1:  LookupTable_ALPT_Markets_PopulateItem i, ALPT_Region_ID.New_England, ALPT_Internal_Market_ID.MA_Central_MA_Westboro, 59
    i = i + 1:  LookupTable_ALPT_Markets_PopulateItem i, ALPT_Region_ID.New_England, ALPT_Internal_Market_ID.MA_Northeastern_MA_Billerica, 58
    i = i + 1:  LookupTable_ALPT_Markets_PopulateItem i, ALPT_Region_ID.New_England, ALPT_Internal_Market_ID.MA_Southeastern_MA_W_Roxbury_2, 57
    i = i + 1:  LookupTable_ALPT_Markets_PopulateItem i, ALPT_Region_ID.New_England, ALPT_Internal_Market_ID.ME_Maine_Hooksett, 62
    i = i + 1:  LookupTable_ALPT_Markets_PopulateItem i, ALPT_Region_ID.New_England, ALPT_Internal_Market_ID.NH_New_Hampshire_Hooksett, 61
    i = i + 1:  LookupTable_ALPT_Markets_PopulateItem i, ALPT_Region_ID.New_England, ALPT_Internal_Market_ID.RI_and_Bristol_County_MA_Taunton, 60
    i = i + 1:  LookupTable_ALPT_Markets_PopulateItem i, ALPT_Region_ID.New_England, ALPT_Internal_Market_ID.VT_and_Western_MA_Counties_Windsor_1, 66
    i = i + 1:  LookupTable_ALPT_Markets_PopulateItem i, ALPT_Region_ID.New_York_Metro, ALPT_Internal_Market_ID.NNJ_Central_Wayne_1, 91
    i = i + 1:  LookupTable_ALPT_Markets_PopulateItem i, ALPT_Region_ID.New_York_Metro, ALPT_Internal_Market_ID.NNJ_East_Jersey_City_2, 87
    i = i + 1:  LookupTable_ALPT_Markets_PopulateItem i, ALPT_Region_ID.New_York_Metro, ALPT_Internal_Market_ID.NNJ_Middlesex_Branchburg_3, 92
    i = i + 1:  LookupTable_ALPT_Markets_PopulateItem i, ALPT_Region_ID.New_York_Metro, ALPT_Internal_Market_ID.NNJ_NorthEast_Jersey_City_1, 86
    i = i + 1:  LookupTable_ALPT_Markets_PopulateItem i, ALPT_Region_ID.New_York_Metro, ALPT_Internal_Market_ID.NNJ_NorthWest_Wayne_2, 90
    i = i + 1:  LookupTable_ALPT_Markets_PopulateItem i, ALPT_Region_ID.New_York_Metro, ALPT_Internal_Market_ID.NNJ_South_Branchburg_2, 89
    i = i + 1:  LookupTable_ALPT_Markets_PopulateItem i, ALPT_Region_ID.New_York_Metro, ALPT_Internal_Market_ID.NNJ_West_Branchburg_1, 88
    i = i + 1:  LookupTable_ALPT_Markets_PopulateItem i, ALPT_Region_ID.New_York_Metro, ALPT_Internal_Market_ID.NYM_Bronx_Upper_Manhattan_WNyack2, 79
    i = i + 1:  LookupTable_ALPT_Markets_PopulateItem i, ALPT_Region_ID.New_York_Metro, ALPT_Internal_Market_ID.NYM_Lower_Manhattan_Mineola2, 84
    i = i + 1:  LookupTable_ALPT_Markets_PopulateItem i, ALPT_Region_ID.New_York_Metro, ALPT_Internal_Market_ID.NYM_Midtown_Manhattan_Yonkers, 80
    i = i + 1:  LookupTable_ALPT_Markets_PopulateItem i, ALPT_Region_ID.New_York_Metro, ALPT_Internal_Market_ID.NYM_Nassau_Farmingdale2, 83
    i = i + 1:  LookupTable_ALPT_Markets_PopulateItem i, ALPT_Region_ID.New_York_Metro, ALPT_Internal_Market_ID.NYM_Queens_Whitestone_2, 82
    i = i + 1:  LookupTable_ALPT_Markets_PopulateItem i, ALPT_Region_ID.New_York_Metro, ALPT_Internal_Market_ID.NYM_Staten_Island_Brooklyn_Whitestone_1, 81
    i = i + 1:  LookupTable_ALPT_Markets_PopulateItem i, ALPT_Region_ID.New_York_Metro, ALPT_Internal_Market_ID.NYM_Suffolk_Farmingdale_1, 85
    i = i + 1:  LookupTable_ALPT_Markets_PopulateItem i, ALPT_Region_ID.New_York_Metro, ALPT_Internal_Market_ID.NYM_Westch_Rockld_Putnam_WNyack1, 78
    i = i + 1:  LookupTable_ALPT_Markets_PopulateItem i, ALPT_Region_ID.Ohio_Pennsylvania, ALPT_Internal_Market_ID.OH_Akron_1, 246
    i = i + 1:  LookupTable_ALPT_Markets_PopulateItem i, ALPT_Region_ID.Ohio_Pennsylvania, ALPT_Internal_Market_ID.OH_Akron_2, 247
    i = i + 1:  LookupTable_ALPT_Markets_PopulateItem i, ALPT_Region_ID.Ohio_Pennsylvania, ALPT_Internal_Market_ID.OH_Cincinnati, 249
    i = i + 1:  LookupTable_ALPT_Markets_PopulateItem i, ALPT_Region_ID.Ohio_Pennsylvania, ALPT_Internal_Market_ID.OH_Cincinnati_Duff_1, 250
    i = i + 1:  LookupTable_ALPT_Markets_PopulateItem i, ALPT_Region_ID.Ohio_Pennsylvania, ALPT_Internal_Market_ID.OH_Cleveland_Cleveland_1, 244
    i = i + 1:  LookupTable_ALPT_Markets_PopulateItem i, ALPT_Region_ID.Ohio_Pennsylvania, ALPT_Internal_Market_ID.OH_Columbus, 248
    i = i + 1:  LookupTable_ALPT_Markets_PopulateItem i, ALPT_Region_ID.Ohio_Pennsylvania, ALPT_Internal_Market_ID.OH_Columbus_Lewis_Center_1, 252
    i = i + 1:  LookupTable_ALPT_Markets_PopulateItem i, ALPT_Region_ID.Ohio_Pennsylvania, ALPT_Internal_Market_ID.OH_Columbus_Lewis_Center_2, 253
    i = i + 1:  LookupTable_ALPT_Markets_PopulateItem i, ALPT_Region_ID.Ohio_Pennsylvania, ALPT_Internal_Market_ID.OH_Dayton_Duff_2, 251
    i = i + 1:  LookupTable_ALPT_Markets_PopulateItem i, ALPT_Region_ID.Ohio_Pennsylvania, ALPT_Internal_Market_ID.OH_NW_Ohio_Maumee, 254
    i = i + 1:  LookupTable_ALPT_Markets_PopulateItem i, ALPT_Region_ID.Ohio_Pennsylvania, ALPT_Internal_Market_ID.OH_Toledo_Cleveland_2, 245
    i = i + 1:  LookupTable_ALPT_Markets_PopulateItem i, ALPT_Region_ID.Ohio_Pennsylvania, ALPT_Internal_Market_ID.PA_Pittsburgh_Bridgeville, 242
    i = i + 1:  LookupTable_ALPT_Markets_PopulateItem i, ALPT_Region_ID.Ohio_Pennsylvania, ALPT_Internal_Market_ID.PA_Pittsburgh_Bridgeville_2, 243
    i = i + 1:  LookupTable_ALPT_Markets_PopulateItem i, ALPT_Region_ID.Ohio_Pennsylvania, ALPT_Internal_Market_ID.PA_Pittsburgh_Johnstown, 241
    i = i + 1:  LookupTable_ALPT_Markets_PopulateItem i, ALPT_Region_ID.Ohio_Pennsylvania, ALPT_Internal_Market_ID.PA_Pittsburgh_Pittsburgh, 240
    i = i + 1:  LookupTable_ALPT_Markets_PopulateItem i, ALPT_Region_ID.Ohio_Pennsylvania, ALPT_Internal_Market_ID.WV_Huntington_St_Clairesville, 255
    i = i + 1:  LookupTable_ALPT_Markets_PopulateItem i, ALPT_Region_ID.Philadelphia_TriState, ALPT_Internal_Market_ID.DE_Wilmington, 102
    i = i + 1:  LookupTable_ALPT_Markets_PopulateItem i, ALPT_Region_ID.Philadelphia_TriState, ALPT_Internal_Market_ID.NJ_Burlington_Maple_Shade, 100
    i = i + 1:  LookupTable_ALPT_Markets_PopulateItem i, ALPT_Region_ID.Philadelphia_TriState, ALPT_Internal_Market_ID.NJ_Vineland_Wilmington_2, 99
    i = i + 1:  LookupTable_ALPT_Markets_PopulateItem i, ALPT_Region_ID.Philadelphia_TriState, ALPT_Internal_Market_ID.PA_Harrisburg_Harrisburg, 97
    i = i + 1:  LookupTable_ALPT_Markets_PopulateItem i, ALPT_Region_ID.Philadelphia_TriState, ALPT_Internal_Market_ID.PA_Lehigh_Valley_Plymouth_Meeting_2, 103
    i = i + 1:  LookupTable_ALPT_Markets_PopulateItem i, ALPT_Region_ID.Philadelphia_TriState, ALPT_Internal_Market_ID.PA_Philadelphia_Philadelphia, 98
    i = i + 1:  LookupTable_ALPT_Markets_PopulateItem i, ALPT_Region_ID.Philadelphia_TriState, ALPT_Internal_Market_ID.PA_Plymouth_Meeting, 101
    i = i + 1:  LookupTable_ALPT_Markets_PopulateItem i, ALPT_Region_ID.Philadelphia_TriState, ALPT_Internal_Market_ID.PA_Scranton_WilkesBarre_Pittston, 96
    i = i + 1:  LookupTable_ALPT_Markets_PopulateItem i, ALPT_Region_ID.South_Central, ALPT_Internal_Market_ID.AR_Greater_Arkansas, 185
    i = i + 1:  LookupTable_ALPT_Markets_PopulateItem i, ALPT_Region_ID.South_Central, ALPT_Internal_Market_ID.AR_LittleRock, 182
    i = i + 1:  LookupTable_ALPT_Markets_PopulateItem i, ALPT_Region_ID.South_Central, ALPT_Internal_Market_ID.AR_Northwest_AR_Fort_Smith, 183
    i = i + 1:  LookupTable_ALPT_Markets_PopulateItem i, ALPT_Region_ID.South_Central, ALPT_Internal_Market_ID.OK_Tulsa, 184
    i = i + 1:  LookupTable_ALPT_Markets_PopulateItem i, ALPT_Region_ID.South_Central, ALPT_Internal_Market_ID.TN_Memphis_WestTenn, 181
    i = i + 1:  LookupTable_ALPT_Markets_PopulateItem i, ALPT_Region_ID.South_LRA, ALPT_Internal_Market_ID.OK_LRA_Cross, 9003
    i = i + 1:  LookupTable_ALPT_Markets_PopulateItem i, ALPT_Region_ID.South_LRA, ALPT_Internal_Market_ID.OK_LRA_Pioneer, 9000
    i = i + 1:  LookupTable_ALPT_Markets_PopulateItem i, ALPT_Region_ID.South_LRA, ALPT_Internal_Market_ID.OK_LRA_Pioneer_2, 9001
    i = i + 1:  LookupTable_ALPT_Markets_PopulateItem i, ALPT_Region_ID.Southwest, ALPT_Internal_Market_ID.AZ_Phoenix_Gilbert, 21
    i = i + 1:  LookupTable_ALPT_Markets_PopulateItem i, ALPT_Region_ID.Southwest, ALPT_Internal_Market_ID.AZ_Phoenix_PHX, 22
    i = i + 1:  LookupTable_ALPT_Markets_PopulateItem i, ALPT_Region_ID.Southwest, ALPT_Internal_Market_ID.AZ_Phoenix_Tempe, 24
    i = i + 1:  LookupTable_ALPT_Markets_PopulateItem i, ALPT_Region_ID.Southwest, ALPT_Internal_Market_ID.AZ_Tucson, 23
    i = i + 1:  LookupTable_ALPT_Markets_PopulateItem i, ALPT_Region_ID.Southwest, ALPT_Internal_Market_ID.NM_Albuquerque, 25
    i = i + 1:  LookupTable_ALPT_Markets_PopulateItem i, ALPT_Region_ID.Southwest, ALPT_Internal_Market_ID.NV_Las_Vegas, 20
    i = i + 1:  LookupTable_ALPT_Markets_PopulateItem i, ALPT_Region_ID.Southwest, ALPT_Internal_Market_ID.TX_El_Paso, 26
    i = i + 1:  LookupTable_ALPT_Markets_PopulateItem i, ALPT_Region_ID.Upstate_New_York, ALPT_Internal_Market_ID.NY_Albany_North_Greenbush_1, 71
    i = i + 1:  LookupTable_ALPT_Markets_PopulateItem i, ALPT_Region_ID.Upstate_New_York, ALPT_Internal_Market_ID.NY_Buffalo, 70
    i = i + 1:  LookupTable_ALPT_Markets_PopulateItem i, ALPT_Region_ID.Upstate_New_York, ALPT_Internal_Market_ID.NY_North_Greenbush_2, 74
    i = i + 1:  LookupTable_ALPT_Markets_PopulateItem i, ALPT_Region_ID.Upstate_New_York, ALPT_Internal_Market_ID.NY_Rochester, 73
    i = i + 1:  LookupTable_ALPT_Markets_PopulateItem i, ALPT_Region_ID.Upstate_New_York, ALPT_Internal_Market_ID.NY_Syracuse, 72
    i = i + 1:  LookupTable_ALPT_Markets_PopulateItem i, ALPT_Region_ID.Washington_Baltimore_Virginia, ALPT_Internal_Market_ID.DC_Washington_Adelphi, 107
    i = i + 1:  LookupTable_ALPT_Markets_PopulateItem i, ALPT_Region_ID.Washington_Baltimore_Virginia, ALPT_Internal_Market_ID.MD_Baltimore_Catonsville, 108
    i = i + 1:  LookupTable_ALPT_Markets_PopulateItem i, ALPT_Region_ID.Washington_Baltimore_Virginia, ALPT_Internal_Market_ID.MD_Frederick_Woodlawn_2, 109
    i = i + 1:  LookupTable_ALPT_Markets_PopulateItem i, ALPT_Region_ID.Washington_Baltimore_Virginia, ALPT_Internal_Market_ID.MD_Salisbury, 110
    i = i + 1:  LookupTable_ALPT_Markets_PopulateItem i, ALPT_Region_ID.Washington_Baltimore_Virginia, ALPT_Internal_Market_ID.MD_Silver_Spring_Annpolis_Junction, 106
    i = i + 1:  LookupTable_ALPT_Markets_PopulateItem i, ALPT_Region_ID.Washington_Baltimore_Virginia, ALPT_Internal_Market_ID.MD_Southern_Maryland_Woodlawn_1, 111
    i = i + 1:  LookupTable_ALPT_Markets_PopulateItem i, ALPT_Region_ID.Washington_Baltimore_Virginia, ALPT_Internal_Market_ID.VA_Chantilly_2, 117
    i = i + 1:  LookupTable_ALPT_Markets_PopulateItem i, ALPT_Region_ID.Washington_Baltimore_Virginia, ALPT_Internal_Market_ID.VA_Charlottesville_Goods_Bridge, 115
    i = i + 1:  LookupTable_ALPT_Markets_PopulateItem i, ALPT_Region_ID.Washington_Baltimore_Virginia, ALPT_Internal_Market_ID.VA_Fairfax_Chantilly, 112
    i = i + 1:  LookupTable_ALPT_Markets_PopulateItem i, ALPT_Region_ID.Washington_Baltimore_Virginia, ALPT_Internal_Market_ID.VA_Norfolk_Lee_Hall, 113
    i = i + 1:  LookupTable_ALPT_Markets_PopulateItem i, ALPT_Region_ID.Washington_Baltimore_Virginia, ALPT_Internal_Market_ID.VA_Richmond, 114
    i = i + 1:  LookupTable_ALPT_Markets_PopulateItem i, ALPT_Region_ID.Washington_Baltimore_Virginia, ALPT_Internal_Market_ID.VA_Roanoke, 116
    


End Sub
' Utility function to quickly populate items in the ALPT lookup table
Private Sub LookupTable_ALPT_Markets_PopulateItem(idx As Integer, alptRegionID As ALPT_Region_ID, alptInternalMarketID As ALPT_Internal_Market_ID, LTE_Market_ID As Integer)
    
    LookupTable_ALPT_Markets(idx).ALPT_RegionID = alptRegionID
    LookupTable_ALPT_Markets(idx).ALPT_InternalMarketID = alptInternalMarketID
    LookupTable_ALPT_Markets(idx).LTE_Market_ID = LTE_Market_ID
    
End Sub

Public Sub ClusterDef_Expand_Test()

    Dim clusterDef As String, clusterDefList() As String
    Dim idx As Long
    
    Dim expectedValues() As Variant
    Dim expectedValuesCount As Long
    Dim i As Long
    
    clusterDef = "10-998-1,2,3; 8-2-998-1,3; 8-2-101-2; 2-100-3; 2-1-1; 3-2; 98100-2; 102100-1,3"
    
    clusterDefList = ClusterDef_Expand(clusterDef, False)
    
    'Debug.Print Join(clusterDefList, vbCrLf)
    
    
    expectedValues = Array( _
        "10-998-1", "10-998-2", "10-998-3", _
        "8-2-998-1", "8-2-998-3", _
        "8-2-101-2", _
        "2-100-3", _
        "2-1-1", _
        "3-2", _
        "098100-2", _
        "102100-1", "102100-3" _
    )
    
    expectedValuesCount = Array_Count(expectedValues)
        
    If expectedValuesCount <> Array_Count(clusterDefList) Then
        Debug.Print "ClusterDef_Expand_Test: Count mismatch"
        Exit Sub
    End If
    
    For i = 0 To expectedValuesCount - 1
        Dim isMatch As Boolean
        
        isMatch = (clusterDefList(i + LBound(clusterDefList)) = expectedValues(i))
        
        Debug.Print clusterDefList(i + LBound(clusterDefList)); Tab(16); expectedValues(i); Tab(32); IIf(isMatch, "OK", "NOT OK")
    Next i
    
End Sub
' Function ClusterDef_Expand
' --------------------------------------
' Converts a cluster definition string to an array with individual cell UIDs and sectors.'
'
' Can be called directly as an array formula directly from excel. If the size of the array (rows x cols) is less than the
' the expanded cluster list, the last column will contain the remaining clusters separated by a ";"
'
' Each item in the cluster list can follow the following pattern: <Cell UID>-<Sector list>
'    Cell UID can be any of the following patterns: (1) <enodeb_id> OR (2) <sys_id>-<ecp/sn>-<cell> OR <ecp/sn>-<cell>
'    Sector list starts with a dash (-) or underscore (_) and is followed by the following pattern: 1 OR 1,2 OR 1,3 OR ...
'
' Example: Expands 000123-1,2,3;99456-1,2;098297 to an array with the following elements:
'   000123-1, 000123-2, 000123-3, 099456-1, 099456-2, 098297-1, 098297-2, 098297-3
'
Public Function ClusterDef_Expand(clusterDefStr As String, Optional removeInvalidClusterItems As Boolean = True, Optional formatSixDigitEnodeID As Boolean = True) As String()

    'Dim clusterDefStr As String:    clusterDefStr = "8-11-100-3;8-11-1004-1,2;11-286-1,2;990860-2;099861" 'Test string
    'Dim removeInvalidClusterItems As Boolean: removeInvalidClusterItems = True
    
    
    Dim retArray() As String
    Dim retArrIdx As Integer
    Dim retArrCount As Integer

    Dim i As Integer, j As Integer, k As Integer
    
    
    Dim clusterStrArr() As String, clusterStrItem As Variant
    Dim sectorStrArr() As String, sectorStrArrItem As Variant
    
    
    
    Dim cell_uid As String, cell_sectors As String
    
    
    clusterStrArr = Split(clusterDefStr, ";")
    
    
    Dim regExp As New VBScript_RegExp_55.regExp
    Dim regExpMatches As VBScript_RegExp_55.MatchCollection
    
    'Dim regExpMatch As VBScript_RegExp_55.Match
    'Dim regExpSubmatches As VBScript_RegExp_55.SubMatches
    'Dim regExpSubmatch As VBScript_RegExp_55.Match
    
    Dim regExpPattern_cell_uid As String
    Dim regExpPattern_sectorList As String
    
    Dim subMatch_cell_uid As String
    Dim subMatch_enode_id As String
    Dim subMatch_sysid_ecp_sn As String
    Dim subMatch_cell_id As String
    Dim subMatch_sectors As String
    
    ' Build regular expression for cell UID and sector list patterns
    '    Cell UID can be any of the following patterns: (1) <enodeb_id> OR (2) <ecp/sn>-<cell>
    '    Sector list is the following pattern: 1 OR 1,2 OR 1,3 OR ... OR (blank)
    '
    ' The submatches is determined by parts of the pattern which are grouped by parentheses (except for a grouping which starts with ?:)
    ' If the cluster string item matches the pattern below, then submatches will contain the following content:
    '   index 0: entire cell_uid (all cell_uid patterns)
    '   index 1: enodeb ID (pattern 1 only)
    '   index 2: sys id (patterns 2 and 3 only)
    '   index 4: cell id (patterns 2 and 3 only)
    '   index 5: sector list (all patterns)
    regExpPattern_cell_uid = "\s*" ' Optional space
    regExpPattern_cell_uid = regExpPattern_cell_uid & "(" ' Start of pattern grouping for cell uid
    regExpPattern_cell_uid = regExpPattern_cell_uid & "(\d{5,6})" ' 5 or 6 digit enodeb ID
    regExpPattern_cell_uid = regExpPattern_cell_uid & "|" ' OR
    regExpPattern_cell_uid = regExpPattern_cell_uid & "(?:([\d\-]*?[^\-])\-)?(\d{1,4})" ' <sys id/ecp/sn)>-<cell (1-4 digits)>
    regExpPattern_cell_uid = regExpPattern_cell_uid & ")" ' End of pattern
    regExpPattern_sectorList = "(?:[\-_](\d(?:,\d){0,5}))" ' Pattern for  sector list suffix. eg. -1,2,3 or _1,2
    
    regExp.Global = True
    regExp.pattern = "^" & regExpPattern_cell_uid & regExpPattern_sectorList & "$"
    
    
    retArrCount = 0
    retArrIdx = 1
    
    For Each clusterStrItem In clusterStrArr
        Set regExpMatches = regExp.Execute(clusterStrItem)
        

        'Debug.Print clusterStrItem & " : " & IIf(regExpMatches.count = 1, "Valid", "Invalid")
        
        If regExpMatches.count = 1 Then ' Pattern has been matched
            subMatch_cell_uid = regExpMatches(0).SubMatches.Item(0)
            subMatch_enode_id = regExpMatches(0).SubMatches.Item(1)
            'subMatch_sys_id = regExpMatches(0).SubMatches.Item(2)
            subMatch_sysid_ecp_sn = regExpMatches(0).SubMatches.Item(2)
            subMatch_cell_id = regExpMatches(0).SubMatches.Item(3)
            subMatch_sectors = regExpMatches(0).SubMatches.Item(4)
            
            If Len(subMatch_enode_id) > 0 Then
                cell_uid = subMatch_enode_id
                cell_uid = Format(CLng(subMatch_enode_id), "000000") ' format as <6 digit enodeb>
            Else
                cell_uid = IIf(Len(subMatch_sysid_ecp_sn) > 0, subMatch_sysid_ecp_sn & "-", "") & subMatch_cell_id
            End If
            
            
             ' If no sector list is provided, assume all sectors - alpha, beta, and gamma - REQUIRED
            'If subMatch_sectors = "" Then subMatch_sectors = "0" '"1,2,3"
             
             ' split sector list
            sectorStrArr = Split(subMatch_sectors, ",")
        
        
            For Each sectorStrArrItem In sectorStrArr
                ' add individual enode B and sector separated by
                retArrCount = retArrCount + 1
                ReDim Preserve retArray(1 To retArrCount) As String
                
                ' format as <6 digit enodeb>-<sector>
                retArray(retArrCount) = cell_uid & "-" & sectorStrArrItem
            Next sectorStrArrItem
            
        Else ' Pattern match failed
            
            If removeInvalidClusterItems = False Then
                retArrCount = retArrCount + 1
                ReDim Preserve retArray(1 To retArrCount) As String
                retArray(retArrCount) = "ClusterDef_Expand: Invalid item " & clusterStrItem
            End If
            
        End If
        
        
        
    Next clusterStrItem
    
    
    Set regExp = Nothing
    Set regExpMatches = Nothing
    
    
    
    ' If calling directly from excel as an array formula, then modify the result to fit the cells
    If IsObject(Application.Caller) Then
        Dim callerRows As Long, callerCols As Long
        Dim isLastCellInCallerArray As Boolean
        Dim newRetArray() As String
    
        callerRows = Application.Caller.Rows.count
        callerCols = Application.Caller.columns.count
        
        retArrIdx = 1
        
        ReDim newRetArray(1 To callerRows, 1 To callerCols) As String
        
        For i = 1 To callerRows
            For j = 1 To callerCols
                isLastCellInCallerArray = IIf(i = callerRows And j = callerCols, True, False)
                
                
                If isLastCellInCallerArray And retArrIdx < retArrCount Then
                    ' Caller array size is less than expanded cluster size
                    ' Fill the last array item with the rest of the cluster list seperated by a semi-colon
                    newRetArray(i, j) = retArray(retArrIdx)
                    
                    For k = retArrIdx + 1 To retArrCount
                        newRetArray(i, j) = newRetArray(i, j) & ";" & retArray(k)
                    Next k
                Else
                    If retArrIdx > retArrCount Then
                        ' Caller array size is greater than expanded cluster size
                        ' Fill array item with empty string
                        newRetArray(i, j) = ""
                    Else
                        newRetArray(i, j) = retArray(retArrIdx)
                        retArrIdx = retArrIdx + 1
                    End If
                End If
                
            Next j
        Next i
        
        ClusterDef_Expand = newRetArray
    Else
        ClusterDef_Expand = retArray
    End If
    
    
End Function

Public Function ClusterDef_Expand_Multi_WIP() As String()

    ' Work-in-progress: Do not use
    
    Dim clusterNames As Variant: clusterNames = Range("Config!$G$32:$G$34")
    Dim clusterDefs As Variant: clusterDefs = Range("Config!$H$32:$H$34")
    
    Dim i As Long, j As Long
    
    Dim clusterNameStr() As String
    Dim clusterDefStr() As String
    
    
    Dim arrDim As Long, arrCount As Long
    Dim dataArr() As Variant
    
    If IsArray(clusterNames) Then
        If Not IsArray(clusterDefs) Then Err.Raise -1, , "Param ClusterDef must also be an array"
        
        If LBound(clusterNames) <> LBound(clusterDefs) Or UBound(clusterNames) <> UBound(clusterDefs) _
            Or Array_NumDimensions(clusterNames) <> Array_NumDimensions(clusterDefs) Then
                Err.Raise -1, , "Param clusterNames/ClusterDef must be similiar arrays"
        End If
        
        
        arrDim = Array_NumDimensions(clusterNames)
        
        ' Convert 2D arrays to 1D arrays
        ReDim clusterNameStr(LBound(clusterNames) To UBound(clusterNames)) As String
        ReDim clusterDefStr(LBound(clusterNames) To UBound(clusterNames)) As String
        
        For i = LBound(clusterNames) To UBound(clusterNames)
            If arrDim = 2 Then
                clusterNameStr(i) = clusterNames(i, 1)
                clusterDefStr(i) = clusterDefs(i, 1)
            Else
                clusterNameStr(i) = clusterNames(i)
                clusterDefStr(i) = clusterDefs(i)
            End If
        Next i
        

    ElseIf TypeName(clusterNames) = "Range" Then
        dataArr = clusterNames
  
        Debug.Print
    Else
        ClusterDef_Expand_Multi_WIP = CVErr(xlErrValue)
    End If


    Dim clusterDef_Expanded() As String
    Dim clusterDefList() As Variant
    Dim clusterDefList_ClusterNames() As Variant



    For i = LBound(clusterNameStr) To UBound(clusterNameStr)
        clusterDef_Expanded = ClusterDef_Expand(clusterDefStr(i), False)
        clusterDefList = Array_AppendMultiple(clusterDefList, clusterDef_Expanded)
        
        clusterDefList_ClusterNames = Array_AppendMultiple(clusterDefList_ClusterNames, Array_CreateAndFill(Array_Count(clusterDef_Expanded), clusterNameStr(i)))
    Next i

                'clusterDef_Expanded = ClusterDef_Expand(SE_ClusterDef_LTE_Event, False)
                'clusterDef_Expanded = ClusterDef_LTE_Add_AWS_eNodeBs(clusterDef_Expanded)
                
                'clusterDefList_LTE = Array_AppendMultiple(clusterDefList_LTE, clusterDef_Expanded)
                'clusterDefList_LTE_ClusterTypes = Array_AppendMultiple(clusterDefList_LTE_ClusterTypes, Array_CreateAndFill(Array_Count(clusterDef_Expanded), "Macro"))


    ' Convert cluster names and expansions to 2xN array
    Dim retArr() As String

    arrCount = Array_Count(clusterDefList)
    
    ReDim retArr(1 To arrCount, 1 To 2) As String
    
    For i = LBound(clusterDefList) To UBound(clusterDefList)
        retArr(i - LBound(clusterDefList) + 1, 1) = clusterDefList_ClusterNames(i)
        retArr(i - LBound(clusterDefList) + 1, 2) = clusterDefList(i)
    Next i
    
    
    If IsObject(Application.Caller) Then
        retArr = Application.WorksheetFunction.transpose(retArr)
    End If
    
    
    ClusterDef_Expand_Multi_WIP = retArr
    


End Function



' Extract only the unique cell UIDs from a cluster definition
' A cell UID can be in the forms of <enodeb_id> or <ecp>-<cell> or <sn>-<cell>
'
' Note: Numbers with leading zeros are considered different from their non-leading zero counterparts. e.g. 0123 and 123 are not considered as different numbers
'
' Input can be an array OR a string separated by semi-colons ";"
Public Function ClusterDef_Extract_UniqueCellList(clusterDefList As Variant, Optional keepOnlyCellID As Boolean = False) As String()


    'Dim clusterDefList As Variant: clusterDefList = "000123-1;000123-2;000123-3;99456-1,2;5-23-1,2;5-6-1,2;098297"
    
    
    Dim clusterDefListArr() As String, clusterDefListArrItem As Variant
    
    Dim retCellList() As String
    Dim retCellListCount As Long
    
    Dim i As Long, j As Long
    Dim lastPos1 As Integer, lastPos2 As Integer, sectorListStartPos As Integer
    
    Dim cell_uid As String
    
    If IsArray(clusterDefList) Then
        clusterDefListArr = Array_ToString(clusterDefList)
        
        If Array_Count(clusterDefListArr) = 0 Then ClusterDef_Extract_UniqueCellList = retCellList: Exit Function
    Else
        
        If IsEmpty(clusterDefList) Then ClusterDef_Extract_UniqueCellList = retCellList: Exit Function
    
        clusterDefListArr = Split(clusterDefList, ";")
    End If
    
    
    
    
    retCellListCount = 0
    
    
    For Each clusterDefListArrItem In clusterDefListArr
        ' Sector list starts at the position of the last _ or - character within the cluster definition
        lastPos1 = InStrRev(clusterDefListArrItem, "-")
        lastPos2 = InStrRev(clusterDefListArrItem, "_")
        
        sectorListStartPos = lastPos1
        If lastPos2 > lastPos1 Then sectorListStartPos = lastPos2
        
        If sectorListStartPos > 0 Then
            cell_uid = Left$(clusterDefListArrItem, sectorListStartPos - 1)
        Else
            cell_uid = clusterDefListArrItem ' no sector list
        End If
        
        If keepOnlyCellID = True Then
            lastPos1 = InStrRev(cell_uid, "-")
            
            If lastPos1 > 0 Then cell_uid = Mid$(cell_uid, lastPos1 + 1)
        End If
        
        ' Try to find cell UID in list. If not in list, then add it
        Dim foundCellInList As Boolean
        
        foundCellInList = False
        
        For i = 1 To retCellListCount
            If retCellList(i) = cell_uid Then
                foundCellInList = True
                Exit For
            End If
        Next i
        
        If foundCellInList = False Then
            retCellListCount = retCellListCount + 1
            ReDim Preserve retCellList(1 To retCellListCount) As String
            retCellList(retCellListCount) = cell_uid
        End If
        
    Next clusterDefListArrItem
    
    
    ClusterDef_Extract_UniqueCellList = retCellList

End Function

' Returns multidimensional array of clusters grouped by market.
' <market>-<cell>-<sector> (Voice/DO) or <market><cell>-<sector> (LTE)
'
' Example 1: Converts clusters 80123-1,2; 90456-1; 100789-3: to
'   Return Value: Array( Array(80123-1,80123-2), Array(90456-1), Array(100789-3)
'   Markets: Array( 080, 090, 100 )
'
' Example 2: Converts clusters 7-123-1,2; 6-456-1; 5-789-3: to
'   Return Value: Array( Array(7-123-1,7-123-2), Array(6-456-1), Array(5-789-3)
'   Markets: Array( 0007, 0006, 0005 )
'
'
' Input can be an array OR a string separated by semi-colons ";"
Public Function ClusterDef_GroupByMarket(clusterDef As Variant, Optional ByRef retMarkets As Variant) As Variant


   ' Dim clusterDef As Variant: clusterDef = "000123-1;000123-2;000123-3;99456-1,2;5-23-1,2;5-6-1,2;3-57-1,2,3;098297"
    
    
    
    Dim clusterDefList() As String
    
    If IsArray(clusterDef) Then
        clusterDefList = clusterDef
    Else
        clusterDefList = ClusterDef_Expand(CStr(clusterDef))
    End If
    
    Debug.Assert UBound(clusterDefList) > 0
    
    
    
    Dim retClusters_MarketCount As Long
    Dim retClusters_Markets() As Variant
    Dim retClusters_CountByMarket() As Long
    
    Dim retClusters As Variant
    
    Dim i As Long, j As Long
    Dim lastPos1 As Integer, lastPos2 As Integer, sectorStartPos As Integer, cellStartPos As Integer
    Dim marketIdx As Integer
    
    Dim cell_sector_part As String, market_part As Variant
    Dim enode_str As String, LTE_Mkt_ID As Long
    
    
    
    Dim clusterDefListItemMarkets() As String
    ReDim clusterDefListItemMarkets(LBound(clusterDefList) To UBound(clusterDefList)) As String
    
    
    retClusters_MarketCount = 0
    
    For i = LBound(clusterDefList) To UBound(clusterDefList)
        ' Sector list starts at the position of the last _ or - character within the cluster definition
        lastPos1 = InStrRev(clusterDefList(i), "-")
        lastPos2 = InStrRev(clusterDefList(i), "_")
        sectorStartPos = lastPos1
        If lastPos2 > lastPos1 Then sectorStartPos = lastPos2
        
        cellStartPos = InStrRev(clusterDefList(i), "-", sectorStartPos - 1) ' Finds the dash before cell ID in format <market/ecp/sn...>-<cell>-<sector>
        
        If cellStartPos > 0 Then
            ' Dash found. Assume format: <market/ecp/sn...>-<cell>-<sector>
            market_part = Left$(clusterDefList(i), cellStartPos - 1)
            
            If IsNumeric(market_part) Then market_part = CLng(market_part)
        Else
            'Dash not found, the format is probably <enodeb>-<sector> (no dash)
            enode_str = Left$(clusterDefList(i), sectorStartPos - 1)
            
            If Not IsNumeric(enode_str) Then Err.Raise -1, , "Item '" & clusterDefList(i) & "'in cluster has invalid format. Expecting: <enodeb>-<sector> OR <market/ecp/sn/...>-<cell>-<sector>"
            
            LTE_Mkt_ID = Int(CDbl(enode_str) / 1000)
            
            market_part = LTE_Mkt_ID 'format$(LTE_Mkt_ID, "0000")
        End If
        
        clusterDefListItemMarkets(i) = market_part
        
        'Debug.Print clusterDefList(i) & " : " & market_part & " (" & cellStartPos & ")"
    
    
        marketIdx = 0 ' default: market group not found in retClusters_Markets list
        
        For j = 1 To retClusters_MarketCount
            If retClusters_Markets(j) = market_part Then
                marketIdx = j
                Exit For
            End If
        Next j
        
        If marketIdx > 0 Then
            ' market found in list. Add to count
            retClusters_CountByMarket(marketIdx) = retClusters_CountByMarket(marketIdx) + 1
        Else
            ' market not found in list. Add market to list and initialize counts.
            retClusters_MarketCount = retClusters_MarketCount + 1
            ReDim Preserve retClusters_Markets(1 To retClusters_MarketCount) As Variant
            ReDim Preserve retClusters_CountByMarket(1 To retClusters_MarketCount) As Long
            
            retClusters_Markets(retClusters_MarketCount) = market_part
            retClusters_CountByMarket(retClusters_MarketCount) = 1
        End If
        
    Next i
    
    
    ' Now add each cluster item to its appropiate market array
    ReDim retClusters(1 To retClusters_MarketCount) As Variant
    
    
    ' Loop through the markets and then each cluster item. Build array of cluster items for that specific market
    For i = 1 To retClusters_MarketCount
        Dim tmpArr() As String
        Dim tmpIdx As Integer
        
        
        ReDim tmpArr(1 To retClusters_CountByMarket(i)) As String
        tmpIdx = 1
        
        For j = LBound(clusterDefList) To UBound(clusterDefList)
            If clusterDefListItemMarkets(j) = retClusters_Markets(i) Then
                tmpArr(tmpIdx) = clusterDefList(j)
                tmpIdx = tmpIdx + 1
            End If
        Next j
        
        retClusters(i) = tmpArr
    Next i
    
    
    ClusterDef_GroupByMarket = retClusters
    
    If Not IsMissing(retMarkets) Then retMarkets = retClusters_Markets

End Function

' Extract only the unique enodeIDs from a cluster definition
'
' Input can be an array OR a string separated by semi-colons ";"
Public Function ClusterDef_Extract_eNodeList(clusterDefList As Variant) As Long()


    'Dim clusterDefList As Variant: clusterDefList = "000123-1;000123-2;000123-3;99456-1,2;5-23-1,2;5-6-1,2;098297"
    
    Dim cellList() As String
    Dim retList() As Long
    Dim retListCount As Long
    
    cellList = ClusterDef_Extract_UniqueCellList(clusterDefList)
    
    If IsEmpty(cellList) Then ClusterDef_Extract_eNodeList = retList: Exit Function
    
    'Dim clusterDefList As Variant: clusterDefList = "000123-1;000123-2;000123-3;99456-1,2;5-23-1,2;5-6-1,2;098297"


    
    retListCount = 0
    
    Dim cellListItem As Variant
    
    
    Dim i As Integer
    
    
    For Each cellListItem In cellList
        ' Disregard any cell uid which is not numeric
        If IsNumeric(cellListItem) Then
            Dim enodeID As Long
            
            enodeID = CLng(cellListItem)
    
            ' Try to find enode ID in list. If not in list, then add it
            Dim foundInList As Boolean
            
            foundInList = False
            
            For i = 1 To retListCount
                If retList(i) = enodeID Then
                    foundInList = True
                    Exit For
                End If
            Next i
            
            If foundInList = False Then
                retListCount = retListCount + 1
                ReDim Preserve retList(1 To retListCount) As Long
                retList(retListCount) = enodeID
            End If
        End If
    Next cellListItem
    
    
    ClusterDef_Extract_eNodeList = retList
    
End Function

' Takes a cluster definition list (e.g. from ClusterDef_Expand) and adds the corresponding AWS eNodeBs. (LTE Market ID + 300)
Public Function ClusterDef_LTE_Add_AWS_eNodeBs(clusterDefList As Variant, Optional filterOnlyValid = False) As String()
        
    ' 2015-05-13 - Fixed sector count bug
    
    
    'Dim clusterDefList As Variant: clusterDefList = ClusterDef_Expand("000123-1,3;99456-2;asdf;098297-1,2,3")
    'Dim filterOnlyValid As Boolean: filterOnlyValid = False
    
    

    If Not IsArray(clusterDefList) Then
        ClusterDef_LTE_Add_AWS_eNodeBs = clusterDefList
        Exit Function
    End If
    
    
    Dim retArray() As String
    Dim retArrIdx As Integer
    Dim retArrCount As Integer
    

    Dim enodeList_IDs() As Long
    Dim enodeList_Sectors() As Integer ' Store sectors in bitwise format (power of 2s) - 1=Alpha,2=Beta,4=Gamma,8=Delta,etc...
    Dim enodeList_SectorCount() As Integer ' Number of sectors for each NB
    Dim enodeListCount As Long
    
    ' Placeholder for input eNBs which are not 700 enBs (2nd or 3rd AWS/PCS LTE carriers)
    Dim enodeCluster_non700() As String
    Dim enodeCluster_non700_count As Integer
    
    ' Placeholder for temporary
    Dim unknownListItems() As String
    Dim unknownListItemCount As Long
    
    Dim i As Integer, j As Integer, k As Integer
    Dim foundInList As Boolean, foundIdx As Integer
    
    Dim clusterDefListItem As Variant
    Dim sectorStrArr() As String, sectorStrArrItem As Variant
    
    Dim cell_uid As String, cell_sectors As String
    
    Dim regExp As New VBScript_RegExp_55.regExp
    Dim regExpMatches As VBScript_RegExp_55.MatchCollection
    
    ' Regular expression for <enodeb_id> with a single sector
    ' Examples: 012345-1 OR 012345_3 OR 12345-2
    regExp.Global = True
    regExp.pattern = "^(\d{5,6})(?:[\-_](\d))$" ' Pattern for eNodeB ID and ONE sector (<enode>-<sector>)
    
    
    ' Begin - Go through each cluster item, determine the base enodeB ID and create list of all eNBs and sectors.
    For Each clusterDefListItem In clusterDefList
        Set regExpMatches = regExp.Execute(clusterDefListItem)
        
        If regExpMatches.count = 1 Then 'Valid match
            cell_uid = regExpMatches(0).SubMatches.Item(0)      ' enode ID
            cell_sectors = regExpMatches(0).SubMatches.Item(1)  ' sectors
            
            Dim enodeB_ID As Long, LTE_Mkt_ID As Long, cell_ID As Long, cell_sector As Integer
            
            
            enodeB_ID = CLng(cell_uid)
            LTE_Mkt_ID = Int(enodeB_ID / 1000) 'Mod 300 ' LTE Market ID of first eNB
            cell_ID = enodeB_ID Mod 1000
            cell_sector = 2 ^ (CInt(cell_sectors) - 1) ' Convert to bitwise format (power of 2s) - 1=Alpha,2=Beta,4=Gamma,8=Delta,etc...
            
            'If enodeB_ID = 102011 Then
            '    Debug.Print
            'End If
            
            If LTE_Mkt_ID > 0 And LTE_Mkt_ID < 300 Then
                ' Is eNB already in list?
                foundIdx = 0
                For i = 1 To enodeListCount
                    If enodeList_IDs(i) = enodeB_ID Then
                        foundIdx = i
                        Exit For
                    End If
                Next i
                
                If foundIdx > 0 Then
                    ' Set sector bit for eNB if not set
                    If Not (enodeList_Sectors(foundIdx) And cell_sector) Then  'bitwise and
                        enodeList_Sectors(foundIdx) = enodeList_Sectors(foundIdx) Or cell_sector
                        enodeList_SectorCount(foundIdx) = enodeList_SectorCount(foundIdx) + 1
                    End If
                Else
                    enodeListCount = enodeListCount + 1
                    ReDim Preserve enodeList_IDs(1 To enodeListCount) As Long
                    ReDim Preserve enodeList_Sectors(1 To enodeListCount) As Integer
                    ReDim Preserve enodeList_SectorCount(1 To enodeListCount) As Integer
                    
                    enodeList_IDs(enodeListCount) = enodeB_ID
                    enodeList_Sectors(enodeListCount) = cell_sector
                    enodeList_SectorCount(enodeListCount) = 1
                End If
            ElseIf LTE_Mkt_ID > 300 Then
                ' eNB has market ID > 300, then IS already second carrier eNB (AWS or PCS-LTE)
                enodeCluster_non700_count = enodeCluster_non700_count + 1
                ReDim Preserve enodeCluster_non700(1 To enodeCluster_non700_count) As String
                enodeCluster_non700(enodeCluster_non700_count) = clusterDefListItem
            Else
                ' Invalid Market ID - assume no corresponding AWS enodeB
                unknownListItemCount = unknownListItemCount + 1
                ReDim Preserve unknownListItems(1 To unknownListItemCount) As String
                unknownListItems(unknownListItemCount) = clusterDefListItem
            End If
        Else
            ' Pattern does not match eNB-sector
            unknownListItemCount = unknownListItemCount + 1
            ReDim Preserve unknownListItems(1 To unknownListItemCount) As String
            unknownListItems(unknownListItemCount) = clusterDefListItem
        End If
       
    Next clusterDefListItem
    ' End - Loop through each cluster item, determine the base enodeB ID and create list of eNBs
    
    
    retArrCount = 0
    retArrIdx = 1
    
    Const MAX_SECTORS_PER_ENB As Integer = 6 ' Assuming 6 sectors/eNB may be possible sometime in the future
    Const POSSIBLE_ENODES_PER_CELL = 2 ' 2 eNBs - (base eNB for 700 + AWS eNB)
    
    Dim sectorBitCheck As Integer, sectorInList As Boolean
    Dim sectorIdx As Integer, offsetIdx As Integer
    
    ' End - Loop through each cluster item, determine the base enodeB ID and create list of eNBs
    For i = 1 To enodeListCount
        ' Pre-allocate array
        retArrCount = retArrCount + (POSSIBLE_ENODES_PER_CELL * enodeList_SectorCount(i))
        ReDim Preserve retArray(1 To retArrCount) As String
        
        
        sectorIdx = 0
        
        For j = 1 To MAX_SECTORS_PER_ENB
            '  bitwise check
            sectorBitCheck = enodeList_Sectors(i) And (2 ^ (j - 1))
            
            
            If sectorBitCheck > 0 Then
                For k = 1 To POSSIBLE_ENODES_PER_CELL
                    ' Calculate offsetIdx to fill array in the following pattern: eNB1-1, eNB1-2, AWS_eNB1-1, AWS_eNB1-2, eNB2-1, AWS_eNB2-1, ...
                    offsetIdx = sectorIdx + (k - 1) * enodeList_SectorCount(i)
                    
                    'Debug.Print Format(enodeList_IDs(i), "000000") & ":" & enodeList_SectorCount(i) & ":" & sectorIdx & ":" & k - 1 & " = " & offsetIdx
                   
                    
                    retArrIdx = retArrCount - (POSSIBLE_ENODES_PER_CELL * enodeList_SectorCount(i)) + offsetIdx + 1
                    
                    
                    enodeB_ID = enodeList_IDs(i)
                    
                    If k = 2 Then enodeB_ID = enodeB_ID + 300000 ' AWS enode - add 300 to LTE_MKT ID
                    ' Need to add another line if POSSIBLE_ENODES_PER_CELL > 2
                    
                    retArray(retArrIdx) = Format(enodeB_ID, "000000") & "-" & j  ' 6 digit enodeB & sector
                Next k

                
                sectorIdx = sectorIdx + 1
            End If
        Next j
    Next i
    
    
    
    ' Add non-700 eNBs back to list
    If enodeCluster_non700_count > 0 Then
        retArrCount = retArrCount + enodeCluster_non700_count
        ReDim Preserve retArray(1 To retArrCount) As String
        
        For i = 1 To enodeCluster_non700_count
            retArrIdx = retArrCount - enodeCluster_non700_count + i
            retArray(retArrIdx) = enodeCluster_non700(i)
        Next i
    End If
    
    ' Add unrecognized cluster items back to list
    If filterOnlyValid = False And unknownListItemCount > 0 Then
        retArrCount = retArrCount + unknownListItemCount
        ReDim Preserve retArray(1 To retArrCount) As String
        
        For i = 1 To unknownListItemCount
            retArrIdx = retArrCount - unknownListItemCount + i
            retArray(retArrIdx) = unknownListItems(i)
        Next i
    End If
    

    ClusterDef_LTE_Add_AWS_eNodeBs = retArray
    
    
End Function


' Sub ClusterDef_FilterWorksheet
' -----------------------------------------------------------------
' Removes rows which are not in a cluster list from a worksheet
'
' Works by creating a row universal ID which is the column values separated by dash (-) and compares with items in the cluster list
' The row universal ID is created from the columns listed in clusterColumnNumbers (array)
' The formatting of each column value is controlled by clusterColumnFormats and via the Format() function. This is useful  if the column value needs
' to be formatted to "align" with the cluster definition. For example, 5 digit enodeB IDs can be formatted into 6-digit enodeB IDs using the format "000000"
Public Sub ClusterDef_FilterWorksheet(ws As Worksheet, clusterDefList() As String, clusterColumnNumbers As Variant, Optional clusterColumnFormats As Variant)


    'Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("MPT")
    'Dim clusterDefList() As String
    
    'clusterDefList = ClusterDef_Expand("319-2; 515-1")
    
    'Dim colNbr_BTS As Integer, colNbr_Sect As Integer
    'colNbr_BTS = Worksheet_GetColumnByName(ws, "BTS")
    'colNbr_Sect = Worksheet_GetColumnByName(ws, "Sector")

    'Debug.Assert colNbr_BTS > 0
   ' Debug.Assert colNbr_Sect > 0
    
    'Dim clusterColumnNumbers As Variant: clusterColumnNumbers = Array(colNbr_BTS, colNbr_Sect)
    'Dim clusterColumnFormats As Variant
    
    
    
 
    
    'If False Then
    If Not IsArray(clusterColumnNumbers) Then Err.Raise -1, , "clusterColumnNumbers must be an array of numbers"  ' cannot do anything here
    
    If Not IsMissing(clusterColumnFormats) And Not IsArray(clusterColumnFormats) Then Err.Raise -1, , "clusterColumnFormats must be an array"

    If Not IsMissing(clusterColumnFormats) Then
        If UBound(clusterColumnFormats) <> UBound(clusterColumnNumbers) Then
            Err.Raise -1, , "clusterColumnFormats array size must be the same as clusterColumnNumbers"
        End If
    End If
    'End If
    

    
    Dim i As Long, j As Long
    Dim colNbr_Last As Long, rowNbr_Last As Long
    
    
    colNbr_Last = Worksheet_GetLastColumn(ws)
    rowNbr_Last = Worksheet_GetLastRow(ws)
    
    
    Dim clusterDefItem As Variant, foundRowInClusterDef As Boolean
    
    Dim dataRng As Range
    Dim data() As Variant
    
    Dim data_keepRowFlag() As Boolean
    Dim filteredData_rowCount As Long
    
    Set dataRng = ws.Cells(1, 1).Resize(rowNbr_Last, colNbr_Last)
    
    data = dataRng.value
    
    ReDim data_keepRowFlag(1 To rowNbr_Last) As Boolean
    
    data_keepRowFlag(1) = True 'header row
    filteredData_rowCount = 1 'header row
    
    
    For i = rowNbr_Last To 2 Step -1
        Dim row_UID As String 'row_UID needs to match the pattern in the cluster definition <cell_id>-<sector>: eg. <enode>-<sector>, <sn>-<cell>-<sector>, etc...
        Dim colNbr As Variant
        
        row_UID = ""
        
        For j = LBound(clusterColumnNumbers) To UBound(clusterColumnNumbers)
            If IsArray(clusterColumnFormats) Then
                If Len(clusterColumnFormats(j)) > 0 Then
                    row_UID = row_UID & Format$(ws.Cells(i, clusterColumnNumbers(j)), clusterColumnFormats(j)) & "-"
                Else
                    row_UID = row_UID & ws.Cells(i, clusterColumnNumbers(j)) & "-"
                End If
            Else
                row_UID = row_UID & ws.Cells(i, clusterColumnNumbers(j)) & "-"
            End If
        Next j
        
        If Len(row_UID) > 0 Then row_UID = Left$(row_UID, Len(row_UID) - 1)
        
        foundRowInClusterDef = False
        
        For Each clusterDefItem In clusterDefList
            If clusterDefItem = row_UID Then
                foundRowInClusterDef = True
                Exit For
            End If
        Next clusterDefItem
        
        
        If foundRowInClusterDef Then
            data_keepRowFlag(i) = True
            filteredData_rowCount = filteredData_rowCount + 1
        Else
            data_keepRowFlag(i) = False
        End If
    Next i
    
    
    
    Dim filteredData() As Variant
    ReDim filteredData(1 To filteredData_rowCount, 1 To colNbr_Last) As Variant
    
    Dim curRow As Long
    curRow = 1
    
    
    For i = 1 To rowNbr_Last
        If data_keepRowFlag(i) Then
            For j = 1 To colNbr_Last
                filteredData(curRow, j) = data(i, j)
            Next j
            
            curRow = curRow + 1
        End If
    Next i
    
    
    'Application.ScreenUpdating = False ' Stop screen painting (speeds up processing)
    'Application.Calculation = xlCalculationManual
    
    dataRng.ClearContents
    dataRng.Resize(filteredData_rowCount, colNbr_Last) = filteredData
    
    'Application.ScreenUpdating = True ' Repaint screen
    'Application.Calculation = xlCalculationAutomatic
    
    
    

End Sub
' Sub ClusterDef_FilterWorksheet
' -----------------------------------------------------------------
' Removes rows which are not in a cluster list from a worksheet
'
' Works by creating a row universal ID which is the column values separated by dash (-) and compares with items in the cluster list
' The row universal ID is created from the columns listed in clusterColumnNumbers (array)
' The formatting of each column value is controlled by clusterColumnFormats and via the Format() function. This is useful  if the column value needs
' to be formatted to "align" with the cluster definition. For example, 5 digit enodeB IDs can be formatted into 6-digit enodeB IDs using the format "000000"
Public Sub ClusterDef_FilterWorksheet_SlowMethod(ws As Worksheet, clusterDefList() As String, clusterColumnNumbers As Variant, Optional clusterColumnFormats As Variant)

    
    
    If Not IsArray(clusterColumnNumbers) Then Err.Raise -1, , "clusterColumnNumbers must be an array of numbers"  ' cannot do anything here
    
    If Not IsMissing(clusterColumnFormats) And Not IsArray(clusterColumnFormats) Then Err.Raise -1, , "clusterColumnFormats must be an array"

    If Not IsMissing(clusterColumnFormats) Then
        If UBound(clusterColumnFormats) <> UBound(clusterColumnNumbers) Then
            Err.Raise -1, , "clusterColumnFormats array size must be the same as clusterColumnNumbers"
        End If
    End If
    

    
    Dim i As Integer, j As Integer
    Dim colNbr_Last As Long, rowNbr_Last As Long
    
    
    colNbr_Last = Worksheet_GetLastColumn(ws)
    rowNbr_Last = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    
    
    Application.ScreenUpdating = False ' Stop screen painting (speeds up processing)
    Application.Calculation = xlCalculationManual
    
    Dim clusterDefItem As Variant, foundRowInClusterDef As Boolean
    
    For i = rowNbr_Last To 2 Step -1
        Dim row_UID As String 'row_UID needs to match the pattern in the cluster definition <cell_id>-<sector>: eg. <enode>-<sector>, <sn>-<cell>-<sector>, etc...
        Dim colNbr As Variant
        
        row_UID = ""
        
        For j = LBound(clusterColumnNumbers) To UBound(clusterColumnNumbers)
            If Not IsMissing(clusterColumnFormats) Then
                If Len(clusterColumnFormats(j)) > 0 Then
                    row_UID = row_UID & Format$(ws.Cells(i, clusterColumnNumbers(j)), clusterColumnFormats(j)) & "-"
                Else
                    row_UID = row_UID & ws.Cells(i, clusterColumnNumbers(j)) & "-"
                End If
            Else
                row_UID = row_UID & ws.Cells(i, clusterColumnNumbers(j)) & "-"
            End If
        Next j
        
        If Len(row_UID) > 0 Then row_UID = Left$(row_UID, Len(row_UID) - 1)
        
        foundRowInClusterDef = False
        
        For Each clusterDefItem In clusterDefList
            If clusterDefItem = row_UID Then
                foundRowInClusterDef = True
                Exit For
            End If
        Next clusterDefItem
        
        ' Row is not in cluster list
        If foundRowInClusterDef = False Then
            ws.Rows(i).EntireRow.Delete
        End If
    Next i
    
    
    
    
    Application.ScreenUpdating = True ' Repaint screen
    Application.Calculation = xlCalculationAutomatic
    
    
    

End Sub

' Replace market section with another (eg. swap ECP with SNs)
Public Function ClusterDef_SubstituteMarketPart(clusterDefList As Variant, findMarkets As Variant, substituteMarkets) As String()

    'Dim clusterDefList As Variant: clusterDefList = ClusterDef_Expand("8-123-1,2; 11-1-1; 99123-2,3")
    'Dim findMarkets As Variant: findMarkets = Array(8, 11)
    'Dim substituteMarkets As Variant: substituteMarkets = Array(54, 52)
    
    If Not IsArray(findMarkets) Then findMarkets = Array(findMarkets)
    If Not IsArray(substituteMarkets) Then substituteMarkets = Array(substituteMarkets)
    
    
    ' size of find markets array need to be equal to array of substitute markets
    Debug.Assert LBound(findMarkets) = LBound(substituteMarkets)
    Debug.Assert UBound(findMarkets) = UBound(substituteMarkets)
    
    Dim findMarketsStr() As String, substituteMarketsStr() As String
    
    findMarketsStr = Array_ToString(findMarkets)
    substituteMarketsStr = Array_ToString(substituteMarkets)
    
    
    Dim clusterDefListArr() As String, clusterDefListArrItem As Variant
    
    
    Dim i As Long, j As Long
    Dim lastPos1 As Integer, lastPos2 As Integer, sectorListStartPos As Integer, cellIdStartPos As Integer
    
    Dim marketPart As String
    Dim cell_sector_uid As String
    
    Dim findIdx As Integer
    
    If IsArray(clusterDefList) Then
        clusterDefListArr = Array_ToString(clusterDefList)
        
        If Array_Count(clusterDefListArr) = 0 Then ClusterDef_SubstituteMarketPart = clusterDefListArr: Exit Function
    Else
        
        If IsEmpty(clusterDefList) Then ClusterDef_SubstituteMarketPart = clusterDefListArr: Exit Function
    
        clusterDefListArr = Split(clusterDefList, ";")
    End If
    
    
    
    'Dim retClusterDefList(LBound(clusterDefListArr) To UBound(clusterDefListArr)) As String
    
    
    For i = LBound(clusterDefListArr) To UBound(clusterDefListArr)
        clusterDefListArrItem = clusterDefListArr(i)
        ' Sector list starts at the position of the last _ or - character within the cluster definition
        lastPos1 = InStrRev(clusterDefListArrItem, "-")
        lastPos2 = InStrRev(clusterDefListArrItem, "_")
        
        sectorListStartPos = lastPos1
        If lastPos2 > lastPos1 Then sectorListStartPos = lastPos2
        
        
        
        If sectorListStartPos > 0 Then
            cellIdStartPos = InStrRev(clusterDefListArrItem, "-", sectorListStartPos - 1)
            
            If cellIdStartPos > 1 Then
                marketPart = Left(clusterDefListArrItem, cellIdStartPos - 1)
                cell_sector_uid = Mid(clusterDefListArrItem, cellIdStartPos + 1)
                
                findIdx = Array_Find(findMarketsStr, marketPart, 1, -1)
                
                If findIdx >= 0 Then
                    clusterDefListArr(i) = substituteMarketsStr(findIdx) & "-" & cell_sector_uid
                End If
            End If
            
        End If
        
        
    Next i
    
    
    
    ClusterDef_SubstituteMarketPart = clusterDefListArr

End Function

Public Function ClusterDef_RemoveMarketPart(clusterDefList As Variant) As String()


    'Dim clusterDefList As Variant: clusterDefList = "000123-1;000123-2;000123-3;99456-1,2;5-23-1,2;5-6-1,2;098297"
    
    Dim clusterDefListArr() As String, clusterDefListArrItem As Variant
    
    'Dim retCellList() As String
    'Dim retCellListCount As Long
    
    Dim i As Long, j As Long
    Dim lastPos1 As Integer, lastPos2 As Integer, sectorListStartPos As Integer
    
    Dim cell_uid As String, cell_sector As String
    
    If IsArray(clusterDefList) Then
        clusterDefListArr = Array_ToString(clusterDefList)
        
        If Array_Count(clusterDefListArr) = 0 Then ClusterDef_RemoveMarketPart = clusterDefListArr: Exit Function
    Else
        
        If IsEmpty(clusterDefList) Then ClusterDef_RemoveMarketPart = clusterDefListArr: Exit Function
    
        clusterDefListArr = Split(clusterDefList, ";")
    End If
    
    
    
    'Dim retClusterDefList(LBound(clusterDefListArr) To UBound(clusterDefListArr)) As String
    
    
    For i = LBound(clusterDefListArr) To UBound(clusterDefListArr)
        clusterDefListArrItem = clusterDefListArr(i)
        ' Sector list starts at the position of the last _ or - character within the cluster definition
        lastPos1 = InStrRev(clusterDefListArrItem, "-")
        lastPos2 = InStrRev(clusterDefListArrItem, "_")
        
        sectorListStartPos = lastPos1
        If lastPos2 > lastPos1 Then sectorListStartPos = lastPos2
        
        If sectorListStartPos > 0 Then
            cell_uid = Left$(clusterDefListArrItem, sectorListStartPos - 1)
            cell_sector = Mid$(clusterDefListArrItem, sectorListStartPos) 'includes dash or underscore
        Else
            cell_uid = clusterDefListArrItem ' no sector list
            cell_sector = ""
        End If
        

        lastPos1 = InStrRev(cell_uid, "-")
            
        If lastPos1 > 0 Then
            clusterDefListArr(i) = Mid$(cell_uid, lastPos1 + 1) & cell_sector
        Else
            clusterDefListArr(i) = cell_uid & cell_sector
        End If
        
    Next i
    
    
    ClusterDef_RemoveMarketPart = clusterDefListArr

End Function


' Extracts text found in between two other pieces of text anywhere in document or other supplied text
' Example: ExtractStringBetweenText("zero-one-two-three-four", "one-", "-three") returns "two"
Public Function ExtractStringBetweenText(textStr As String, prefixStr As String, suffixStr As String, Optional searchInReverse As Boolean = False) As String


    Dim prefixStrPos As Long, suffixStrPos As Long
    
    prefixStrPos = 1
    suffixStrPos = Len(textStr)
    
    If searchInReverse Then
        If Len(suffixStr) > 0 Then suffixStrPos = InStrRev(textStr, suffixStr)
        If Len(prefixStr) > 0 Then prefixStrPos = InStrRev(textStr, prefixStr, suffixStrPos)
    Else
        If Len(prefixStr) > 0 Then prefixStrPos = InStr(1, textStr, prefixStr)
        If Len(suffixStr) > 0 Then suffixStrPos = InStr(prefixStrPos + Len(suffixStr), textStr, suffixStr)
    End If
    
    
    If prefixStrPos = 0 Or suffixStrPos = 0 Then
        ' cannot find prefix or suffix
        ExtractStringBetweenText = ""
        Exit Function
    End If
    
    ExtractStringBetweenText = Mid$(textStr, prefixStrPos + Len(prefixStr), suffixStrPos - prefixStrPos - Len(prefixStr))
    

End Function

Private Function Prompt_UserPass(ByRef user As String, ByRef pass As String) As Boolean
    


    user = InputBox("Enter username (USWIN):", , Environ$("username"))
    
    If Len(user) = 0 Then Prompt_UserPass = False: Exit Function
    
    'pass = InputBoxMasked("Enter password:", "")
    pass = InputBox("Enter password:", "")
    
    If Len(pass) = 0 Then Prompt_UserPass = False: Exit Function
    
    
    Prompt_UserPass = True

End Function

Public Function ALPT_NewSession(Optional user As String, Optional pass As String) As ALPT_Session
    
    Dim session As ALPT_Session
    
    If user = "" Or pass = "" Then
        If Prompt_UserPass(session.user, session.pass) = False Then
            Err.Raise -1, , "User/Pass required"
        End If
    Else
        session.user = user
        session.pass = pass
    End If
    
    session.BASE_URI = "http://vaculpt-vip.nss.vzwnet.com:8282/alte/"
    
    
    Set session.httpReq = New WinHttp.WinHttpRequest 'CreateObject("WinHttp.WinHttpRequest.5.1")

    ALPT_NewSession = session

End Function
Public Sub ALPT_Session_End(session As ALPT_Session)

    Set session.httpReq = Nothing

End Sub

Private Function ALPT_Session_Login(session As ALPT_Session) As Boolean



    If Not session.loggedIn Then
        
        Dim postData As String
        
        postData = URL_BuildQueryString("name", session.user, "password", session.pass, "webServer", session.BASE_URI, "chkUserPref", "on", "sub", "  Login  ")
        
        
        
        'session.httpReq.Option(WinHttpRequestOption_EnableRedirects) = False
        'session.httpReq.Option(WinHttpRequestOption_UserAgentString) = "Mozilla/5.0 (Windows NT 6.1; WOW64; rv:31.0) Gecko/20100101 Firefox/31.0)"
 
        session.httpReq.Open "POST", session.BASE_URI & "login.htm", False
        session.httpReq.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
        session.httpReq.Send postData
        
        
        
        If session.httpReq.Status <> 200 Then ' HTTP_OK = 200
            Err.Raise -1, , "ALPT Login Failed: Unexpected server response code: " & session.httpReq.Status & " (" & session.httpReq.StatusText & ")"
            ALPT_Session_Login = False
            Exit Function
        End If
        
        If InStr(1, session.httpReq.ResponseText, "Login failed") > 0 Then
            Dim errReason As String
            
            errReason = ExtractStringBetweenText(session.httpReq.ResponseText, "<td class=""error"">", "</td>")
            
            Err.Raise -1, , "ALPT Login Failed: " & errReason
            ALPT_Session_Login = False
            Exit Function
        End If
    
        session.loggedIn = True
    End If
    
    
    ALPT_Session_Login = True
    
    

End Function
' ALPT_StructuredAdHocQuery:
' -----------------------------------------------------------------------------------
' Executes ALPT Structured Ad-Hoc query with RAW data and returns the dataResouceName and report ID
'
' queryData must be a url-encoded query string (build with URL_BuildQueryString)
Public Function ALPT_SAH_Query_RAW(session As ALPT_Session, queryData As String) As ALPT_SAH_Result

    Dim result As ALPT_SAH_Result
    
    result.dataResourceName = ""
    result.reportID = 0
    
    

    ALPT_Session_Login session
    
    
    ' Visit structured ad hoc page. This is somehow necessary for ALPT, otherwise the query fails.
    session.httpReq.Open "GET", session.BASE_URI & "sah.htm", False
    session.httpReq.Send
    
    
    ' Add to query data: Exclude Totals, Report Mode=table, Action=Submit
    queryData = queryData & "&" & URL_BuildQueryString( _
        "cbxExcludeTotals", "Y", _
        "radRptMode", "ta", _
        "hidRequestedAction", "submit" _
    )
    
        
    session.httpReq.SetTimeouts 0, -1, -1, -1 ' unlimited timeouts (this will take awhile)
    session.httpReq.Open "POST", session.BASE_URI & "sah.htm", False
    session.httpReq.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    session.httpReq.Send queryData
    
    
    ' Extract data resource name (web filename of XLS data) and SAH report ID
    Dim reportIdStr As String
    
    reportIdStr = ExtractStringBetweenText(session.httpReq.ResponseText, "<input type ='hidden' name='reportId' value='", "'>")
            
    result.dataResourceName = ExtractStringBetweenText(session.httpReq.ResponseText, "<input type ='hidden' name='dataResourceName' value='", "'>")
    result.reportID = CLng(reportIdStr)
    
    
    
    Debug.Print "ALPT Report: " & result.dataResourceName
    
    Debug.Print "ALPT Chart: " & session.BASE_URI & "applets/lte/jChartApplet.jsp?" & _
        URL_BuildQueryString("dataResourceName", result.dataResourceName, "typestart", 1, "fitstart", 2)
    
    
    ALPT_SAH_Query_RAW = result
    

End Function
Public Function ALPT_SAH_Query_Cluster(destWorksheet As String, session As ALPT_Session, clusterDef As Variant, _
    reportLevel As ALPT_SAH_Report_Level, reportType As ALPT_SAH_Report_Type, reportContent As Variant, _
    dateEnd As Date, numberOfDaysToTrend As Integer, _
    Optional reportCarriers As Variant, Optional reportDaysOfWeek As Variant, Optional reportHours As Variant) As ALPT_SAH_Result
        
        Debug.Assert dateEnd < Now()
        Debug.Assert numberOfDaysToTrend > 0
        
        If reportLevel <> ALPT_SAH_Report_Level.euTrancell And reportLevel <> ALPT_SAH_Report_Level.Sector_Carrier Then
            Err.Raise -1, , "Report Level for ALPT_SAH_Query_Cluster must be on the sector level (EUTranCell or Sector-Carrier)"
        End If
        
        
        
        Dim clusterDefList() As String
        
        ' If cluster definition is not already an expanded array, then expand it
        If Not IsArray(clusterDef) Then
            clusterDefList = ClusterDef_Expand(CStr(clusterDef))
        Else
            clusterDefList = Array_ToString(clusterDef)
        End If
        
        
        Dim reportEuTranCell As Variant
        Dim reportDay As Variant
        
        ' Set defaults if not already provided
        If IsMissing(reportCarriers) Then reportCarriers = Array(1, 2, 3, 4) ' All carriers
        If IsMissing(reportDaysOfWeek) Then reportDaysOfWeek = Array(1, 2, 3, 4, 5, 6, 7)  ' All days of the week
        If IsMissing(reportHours) Then reportHours = "All"
    
        reportEuTranCell = Array(1, 2, 3, 4, 5, 6) ' Always include all sectors
       

        
        
        ' TODO: Need to somehow validate if ReportLevel, reportType and ReportContent are a valid combination


        ' Convert eNode IDs into ALPT Internal Cell IDs
        Dim enodeList() As Long, alptCellIds() As Long

    
        enodeList = ClusterDef_Extract_eNodeList(clusterDefList)
        alptCellIds = ALPT_eNodeList_To_CellIDs(session, enodeList)
        
    
        ' Build SAH query params
        Dim postData As String
        
        postData = URL_BuildQueryString( _
            "selReportLevel", reportLevel, _
            "selReportType", reportType, _
            "selContent", reportContent, _
            "selCellSite", alptCellIds, _
            "selEUTranCell", reportEuTranCell, _
            "selCarrier", reportCarriers, _
            "selDay", reportDaysOfWeek, _
            "selHour", reportHours, _
            "selEndDate", Format(dateEnd, "mm/dd/yyyy"), _
            "selDaysToTrend", numberOfDaysToTrend, _
            "cbxExcludeTotals", "Y" _
        )
        
        
        
        
        Dim alptSahQueryResult As ALPT_SAH_Result
        
        alptSahQueryResult = ALPT_SAH_Query_RAW(session, postData)
        
        
        GetWebData_CSV session.BASE_URI & "createcsv?" & URL_BuildQueryString("fileName", alptSahQueryResult.dataResourceName), destWorksheet
        

        
        ' Remove cell/sectors not in cluster list
        Dim colNbr_ENODEB As Integer
        Dim colNbr_EUTRANCELL As Integer
        
        
        Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(destWorksheet)
        colNbr_ENODEB = Worksheet_GetColumnByName(ws, "ENODEB")
        colNbr_EUTRANCELL = Worksheet_GetColumnByName(ws, "EUTRANCELL")
        
        Debug.Assert colNbr_ENODEB > 0
        Debug.Assert colNbr_EUTRANCELL > 0
        
        Dim clusterColumnNumbers As Variant: clusterColumnNumbers = Array(colNbr_ENODEB, colNbr_EUTRANCELL)
        Dim clusterColumnFormats As Variant:  clusterColumnFormats = Array("000000", "0")
        
        
        
        ClusterDef_FilterWorksheet ws, clusterDefList, clusterColumnNumbers, clusterColumnFormats
    
    
        ALPT_SAH_Query_Cluster = alptSahQueryResult
    

End Function





' Calls ALPT getElementList Web service with queryString containing the parameters
' Elements are returned to the alptElementList array.
Private Function ALPT_GetElementList(session As ALPT_Session, queryString As String, ByRef alptElementList() As ALPT_ElementListItem) As Long

    ALPT_Session_Login session
    
    Dim fullURL As String
    
    fullURL = session.BASE_URI & "getElementList.htm?" & queryString
    
    session.httpReq.Open "GET", fullURL, False
    session.httpReq.Send
    
    
    Dim elementCount As Long
    
    
    
    ' Parse XML Document returned by web server. Format is as follows:
    '
    ' Example XML Doc Format:
    ' <elements>
    '   <element value="3371130" description="096011_WAPWALLOPEN" />
    '   <element value="1021202" description="096013_CARBONDALE" />
    '   <element value="3371110" description="096014_MISERICORDIA" />
    '   ...
    ' </elements>
    
    Dim xmlDoc As Object 'MSXML2.DOMDocument
    Dim xmlNodes As Object ' MSXML2.IXMLDOMNodeList
    Dim xmlNodeElement As Object 'MSXML2.IXMLDOMNode

    
    Dim i As Long
   
    Set xmlDoc = CreateObject("MSXML2.DOMDocument") 'New MSXML2.DOMDocument
   
    xmlDoc.LoadXML session.httpReq.ResponseText
   
    Set xmlNodes = xmlDoc.DocumentElement.getElementsByTagName("element")
    
    elementCount = xmlNodes.length
    
    ReDim alptElementList(1 To elementCount) As ALPT_ElementListItem
    
    i = 1
    
    For Each xmlNodeElement In xmlNodes
        alptElementList(i).value = xmlNodeElement.Attributes(0).text
        alptElementList(i).description = xmlNodeElement.Attributes(1).text
        i = i + 1
    Next xmlNodeElement
    

    ALPT_GetElementList = elementCount

    Set xmlDoc = Nothing
    Set xmlNodes = Nothing
    Set xmlNodeElement = Nothing

End Function






Public Function NCWS_CellGroupReport_Cluster(destWorksheet As String, techType As String, cellGroup As String, clusterDef As Variant, _
    reportType As String, reportGroupBy As String, reportContent As Variant, _
    dateEnd As Date, numberOfDaysToTrend As Integer, _U
    Optional reportCarriers As Variant, Optional reportDays As Variant, Optional reportHours As Variant)
    
    
    ' If cluster definition is not already an expanded array, then expand it
    Dim clusterDefList() As String
    Dim cellList As Variant
    
    
    If IsArray(clusterDef) Then
        clusterDefList = Array_ToString(clusterDef)
        cellList = ClusterDef_Extract_UniqueCellList(clusterDef)
    Else
        If IsMissing(clusterDef) Or clusterDef = "" Then
            Err.Raise "Cluster definition is a required parameter"
        Else
            clusterDefList = ClusterDef_Expand(CStr(clusterDef))
            cellList = ClusterDef_Extract_UniqueCellList(clusterDefList)
        End If
    End If
    
    NCWS_CellGroupReport destWorksheet:=destWorksheet, techType:=techType, cellGroup:=cellGroup, _
        cellList:=cellList, _
        reportType:=reportType, reportGroupBy:=reportGroupBy, reportContent:=reportContent, _
        dateEnd:=dateEnd, numberOfDaysToTrend:=numberOfDaysToTrend, _
        reportCarriers:=reportCarriers, reportDays:=reportDays, reportHours:=reportHours
    
    

    ' Remove cell/sectors not in cluster list
    Dim colNbr_SysID As Integer
    Dim colNbr_SN As Integer
    Dim colNbr_ECP As Integer
    Dim colNbr_BTS As Integer
    Dim colNbr_Sect As Integer
    Dim colNbr_Cell As Integer
    
    
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(destWorksheet)
    Dim clusterColumnNumbers As Variant
    
    
    colNbr_SysID = Worksheet_GetColumnByName(ws, "SysID")
    colNbr_Sect = Worksheet_GetColumnByName(ws, "Sect")
    
    Debug.Assert colNbr_SysID > 0
    Debug.Assert colNbr_Sect > 0
    
    If techType = "DO" Then
        colNbr_SN = Worksheet_GetColumnByName(ws, "SN")
        colNbr_BTS = Worksheet_GetColumnByName(ws, "BTS")
    
        Debug.Assert colNbr_SN > 0
        Debug.Assert colNbr_BTS > 0
    
        clusterColumnNumbers = Array(colNbr_SysID, colNbr_SN, colNbr_BTS, colNbr_Sect)
    ElseIf techType = "1X" Then
        colNbr_ECP = Worksheet_GetColumnByName(ws, "ECP")
        colNbr_Cell = Worksheet_GetColumnByName(ws, "Cell")
    
        Debug.Assert colNbr_ECP > 0
        Debug.Assert colNbr_Cell > 0
    
        clusterColumnNumbers = Array(colNbr_SysID, colNbr_ECP, colNbr_Cell, colNbr_Sect)
    End If

        
    ClusterDef_FilterWorksheet ws, clusterDefList, clusterColumnNumbers

End Function
Public Function NCWS_CellGroupReport(destWorksheet As String, techType As String, cellGroup As String, cellList As Variant, _
    reportType As String, reportGroupBy As String, reportContent As Variant, _
    dateEnd As Date, numberOfDaysToTrend As Integer, _
    Optional reportCarriers As Variant, Optional reportDays As Variant, Optional reportHours As Variant)
        
        
        Debug.Assert dateEnd < Now()
        Debug.Assert numberOfDaysToTrend > 0
        Debug.Assert techType = "1X" Or techType = "DO"
        
        
        
        
        
        Dim i As Integer
        Dim reportSectors As Variant
        Dim reportHoursStr() As String
        
        ' Set defaults if not already provided
        If Not IsArray(cellList) Then
            If cellList = "" Then cellList = "all"
        End If
        
        If IsMissing(reportCarriers) Then reportCarriers = Array(1, 2, 3, 4, 5, 6) ' All carriers
        If IsMissing(reportDays) Then reportDays = Array(1, 2, 3, 4, 5, 6, 7) ' All days of the week
        If IsMissing(reportHours) Then reportHours = Array(0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23)
        
        reportSectors = Array(1, 2, 3, 4, 5, 6)     ' Always include all sectors
        
        ' Transform reports from integers (1,2,3,4...) to two character strings (00, 01, 02, ...)
        reportHoursStr = Array_ToString(reportHours, "00")
    
        
        'Dim reportFields As Variant: reportFields = Array("evdousageccs", "userfcart", "fwddatavolumembytes", "userconnatts")
    

        ' Build query params
        Const BASE_URI = "https://ncws.vh.eng.vzwcorp.com:8080"
        
        Dim queryData As String
        Dim myURL As String
        
        
        
        If techType = "DO" Then
            '
            queryData = URL_BuildQueryString("cg", cellGroup, _
                "emon", Format$(dateEnd, "mm"), "eday", Format$(dateEnd, "dd"), "eyr", Format$(dateEnd, "yyyy"), _
                "special_date", 0, _
                "trend", Format$(numberOfDaysToTrend, "00"), _
                "dw", reportDays, _
                "c", cellList, _
                "seclist", "", _
                "username", Environ$("username"), _
                "fms", "all", _
                "hr", reportHoursStr, _
                "ant", reportSectors, _
                "carr", "all", "freq", "all", _
                "ntype", "eq", "n", "0", "threshold", "none", "throp", "gt", "thvalue", "0", "th2", "", "andor2", "and", "threshold2", "none", "throp2", "gt", "thvalue2", "0", "t2_th2", "", "andor3", "and", "threshold3", "none", "throp3", "gt", "thvalue3", "0", "t3_th2", "", _
                "maxrows", "128000", "newwin", "on", "namein", "in", "namefil", "", _
                "report", reportType, "gby", reportGroupBy, "inc", reportContent, _
                "fms", "all", _
                "order", "dt", "evorddir", "Asc", "ntype", "eq", "n", "0", _
                "suphtml", "on" _
            )
            
            '"report", reportType, "gby", reportGroupBy, "fields", reportFields, _


            
            myURL = BASE_URI & "/cgi-bin/ev_reports.pl?" & queryData
        ElseIf techType = "1X" Then
            queryData = URL_BuildQueryString("cg", cellGroup, _
                "emon", Format$(dateEnd, "mm"), "eday", Format$(dateEnd, "dd"), "eyr", Format$(dateEnd, "yyyy"), _
                "special_date", 0, _
                "trend", Format$(numberOfDaysToTrend, "00"), _
                "dw", reportDays, _
                "c", cellList, _
                "seclist", "", _
                "username", Environ$("username"), _
                "fms", "all", _
                "hr", reportHoursStr, _
                "ant", reportSectors, _
                "carr", "all", "freq", "all", _
                "ntype", "eq", "n", "0", "threshold", "none", "throp", "gt", "thvalue", "0", "th2", "", "andor2", "and", "threshold2", "none", "throp2", "gt", "thvalue2", "0", "t2_th2", "", "andor3", "and", "threshold3", "none", "throp3", "gt", "thvalue3", "0", "t3_th2", "", _
                "maxrows", "128000", "newwin", "on", "namein", "in", "namefil", "", _
                "cgreport", reportType, "gbycg", reportGroupBy, "cginc", reportContent, _
                "ordercg", "dt", "orddir", "Asc", _
                "suphtml", "on" _
            )
            
            
            myURL = BASE_URI & "/cgi-bin/cg_reports.pl?" & queryData
        End If
        
        
        Debug.Print "NCWS " & techType & ": " & myURL

        Dim httpReq As New WinHttp.WinHttpRequest
        
        
        httpReq.SetTimeouts 0, -1, -1, -1 ' Unlimited
        'httpReq.SetTimeouts 0, 60000, 30000, 60000 ' Unlimited, 60s (default), 30s (default), 60s (2 x default)
        httpReq.Open "GET", myURL, False
        httpReq.Send
        
    
    
        Dim dataExcelPath As String, dataCsvPath As String
        dataExcelPath = ExtractStringBetweenText(httpReq.ResponseText, "<a href='", "'>Excel Import</a>", True)
        dataCsvPath = ExtractStringBetweenText(httpReq.ResponseText, "<a href='", "'>CSV Output</a>", True)
        
        Set httpReq = Nothing
        
        
        
        myURL = BASE_URI & dataCsvPath
        
        Debug.Print myURL
        
    
        GetWebData_CSV myURL, destWorksheet
        
        
        
        

    
    
    

End Function

Public Function MPT_CellGroupReport_Cluster(destWorksheet As String, techType As String, market As String, clusterDef As Variant, _
    reportType As String, reportGroupBy As String, reportContent As Variant, _
    dateEnd As Date, numberOfDaysToTrend As Integer, _
    Optional reportCarriers As Variant, Optional reportDays As Variant, Optional reportHours As Variant)
        
        
    ' If cluster definition is not already an expanded array, then expand it
    Dim clusterDefList() As String
    Dim cellList As Variant
    
    
    If IsArray(clusterDef) Then
        clusterDefList = Array_ToString(clusterDef)
        cellList = ClusterDef_Extract_UniqueCellList(clusterDef, True)
    Else
        If IsMissing(clusterDef) Or clusterDef = "" Then
            Err.Raise "Cluster definition is a required parameter"
        Else
            clusterDefList = ClusterDef_Expand(CStr(clusterDef))
            cellList = ClusterDef_Extract_UniqueCellList(clusterDefList, True)
        End If
    End If
    
    
    MPT_CellGroupReport destWorksheet:=destWorksheet, techType:=techType, market:=market, cellList:=cellList, _
        reportType:=reportType, reportGroupBy:=reportGroupBy, reportContent:=reportContent, _
        dateEnd:=dateEnd, numberOfDaysToTrend:=numberOfDaysToTrend, _
        reportCarriers:=reportCarriers, reportDays:=reportDays, reportHours:=reportHours
    
        
    ' Remove cell/sectors not in cluster list
    Dim colNbr_SysID As Integer
    Dim colNbr_SN As Integer
    Dim colNbr_BTS As Integer
    Dim colNbr_Sect As Integer
    
    
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(destWorksheet)
    colNbr_BTS = Worksheet_GetColumnByName(ws, "BTS")
    
    If techType = "1X" Then
        colNbr_Sect = Worksheet_GetColumnByName(ws, "Sec")
    ElseIf techType = "DO" Then
        colNbr_Sect = Worksheet_GetColumnByName(ws, "Sector")
    End If
    
    Debug.Assert colNbr_BTS > 0
    Debug.Assert colNbr_Sect > 0
    
    
    
    Dim clusterColumnNumbers As Variant: clusterColumnNumbers = Array(colNbr_BTS, colNbr_Sect)
    
    ClusterDef_FilterWorksheet ws, clusterDefList, clusterColumnNumbers
    
    
        

End Function
' ALPT_CustomReport_Cluster: Runs a Custom ALPT report at the sector level and removes sectors which are not in the cluster
Public Sub ALPT_CustomReport_Cluster(destSheet As String, clusterDef As Variant, dateEnd As Date, daysToTrend As Integer, reportOwner As String, reportTemplate As String, Optional reportType As String = "Hourly Totals", Optional reportLevel = "CARRIER")


    ' Assert invalid params - should have been checked before we got to this point
    Debug.Assert dateEnd <= Now()
    Debug.Assert daysToTrend > 0

    Dim enodeList() As Long
    Dim clusterDefList() As String
    
    
    Dim allowedReportLevels() As Variant
    
    allowedReportLevels = Array("EUTRANCELL", "CARRIER")
    
    If Array_Find(allowedReportLevels, reportLevel) = -1 Then
        Err.Raise -1, , "ALPT_CustomReport_Cluster: Invalid report level: " & reportLevel & vbCrLf & vbCrLf & "Allowed report levels:" & Join(allowedReportLevels, ", ")
    End If
    
    
    
    ' If cluster definition is not already an expanded array, then expand it
    If Not IsArray(clusterDef) Then
        clusterDefList = ClusterDef_Expand(CStr(clusterDef))
    Else
        clusterDefList = Array_ToString(clusterDef)
    End If
    
        
    enodeList = ClusterDef_Extract_eNodeList(clusterDefList)
        


    Dim myURL As String, queryData As String
    

    
    queryData = URL_BuildQueryString( _
        "action", "exectmpl", _
        "tmpl_rpt", reportOwner & "|||" & reportTemplate, _
        "user", Environ$("username"), _
        "rpttype", reportType, _
        "rptlevel", reportLevel, _
        "edate", Format$(dateEnd, "yyyy-mm-dd"), _
        "num_days", daysToTrend, _
        "enodeb", Array_ToString(enodeList, "000000") _
    )
    


    myURL = "http://alpt.vh.eng.vzwcorp.com:8282/alte/reportWebService.htm?" & queryData

    Debug.Print myURL
    

    GetWebData_CSV myURL, destSheet
    
    If ThisWorkbook.Sheets(destSheet).Cells(1, 1) Like "*Unable to pull report*" Then
        Err.Raise -1, , "Problem running ALPT report '" & reportTemplate & "'. No data returned"
    End If
    
    
    ' Remove cell/sectors not in cluster list
    Dim colNbr_ENODEB As Integer
    Dim colNbr_EUTRANCELL As Integer
    
    
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(destSheet)
    colNbr_ENODEB = Worksheet_GetColumnByName(ws, "ENODEB")
    colNbr_EUTRANCELL = Worksheet_GetColumnByName(ws, "EUTRANCELL")
    
    
    
    Debug.Assert colNbr_ENODEB > 0
    
    If colNbr_EUTRANCELL > 0 Then
        Dim clusterColumnNumbers As Variant: clusterColumnNumbers = Array(colNbr_ENODEB, colNbr_EUTRANCELL)
        Dim clusterColumnFormats As Variant:  clusterColumnFormats = Array("000000", "0")
        
        
        
        ClusterDef_FilterWorksheet ws, clusterDefList, clusterColumnNumbers, clusterColumnFormats
    End If


End Sub
' ALPT_CustomReport_Cluster: Runs a Custom ALPT report at the enodeB level
Public Sub ALPT_CustomReport_eNodeList(destSheet As String, enodeList() As Long, dateEnd As Date, daysToTrend As Integer, reportOwner As String, reportTemplate As String, Optional reportType As String = "Hourly Totals", Optional reportLevel = "ENODEB_CARRIER")

    ' Assert invalid params - should have been checked before we got to this point
    Debug.Assert dateEnd <= Now()
    Debug.Assert daysToTrend > 0

    
    
    Dim allowedReportLevels() As Variant
    
    allowedReportLevels = Array("ENODEB", "ENODEB_CARRIER", "ENODEB_HANDSET")
    
    If Array_Find(allowedReportLevels, reportLevel) = -1 Then
        Err.Raise -1, , "ALPT_CustomReport_eNodeList: Invalid report level: " & reportLevel & vbCrLf & vbCrLf & "Allowed report levels:" & Join(allowedReportLevels, ", ")
    End If
    

    Dim myURL As String, queryData As String

    
    queryData = URL_BuildQueryString( _
        "action", "exectmpl", _
        "tmpl_rpt", reportOwner & "|||" & reportTemplate, _
        "user", Environ$("username"), _
        "rpttype", reportType, _
        "rptlevel", reportLevel, _
        "edate", Format$(dateEnd, "yyyy-mm-dd"), _
        "num_days", daysToTrend, _
        "enodeb", Array_ToString(enodeList, "000000") _
    )
    


    myURL = "http://alpt.vh.eng.vzwcorp.com:8282/alte/reportWebService.htm?" & queryData

    Debug.Print myURL
    

    GetWebData_CSV myURL, destSheet


    If ThisWorkbook.Sheets(destSheet).Cells(1, 1) Like "*Unable to pull report*" Then
        Err.Raise -1, , "Problem running ALPT report '" & reportTemplate & "'. No data returned"
    End If


End Sub



Public Function MPT_CellGroupReport(destWorksheet As String, techType As String, market As String, cellList As Variant, _
    reportType As String, reportGroupBy As String, reportContent As Variant, _
    dateEnd As Date, numberOfDaysToTrend As Integer, _
    Optional reportCarriers As Variant, Optional reportDays As Variant, Optional reportHours As Variant)
        
        
        Debug.Assert dateEnd < Now()
        Debug.Assert numberOfDaysToTrend > 0
        Debug.Assert techType = "1X" Or techType = "DO"
        
        
        Const BASE_URI = "http://mpt.vh.eng.vzwcorp.com:8080"
        
        
        
        Dim i As Integer
        Dim reportSectors As Variant
        Dim reportHoursStr() As String
        
        ' Set defaults if not already provided
        If Not IsArray(cellList) Then
            If cellList = "" Then cellList = "all"
        End If
        
        If IsMissing(reportCarriers) Then reportCarriers = Array(1, 2, 3, 4, 5, 6) ' All carriers
        If IsMissing(reportDays) Then reportDays = Array(1, 2, 3, 4, 5, 6, 7) ' All days of the week
        If IsMissing(reportHours) Then reportHours = Array(0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23)
        
        reportSectors = Array(1, 2, 3, 4, 5, 6)     ' Always include all sectors
        
        ' Transform reports from integers (1,2,3,4...) to two character strings (00, 01, 02, ...)
        reportHoursStr = Array_ToString(reportHours, "00")
        'ReDim reportHoursStr(LBound(reportHours) To UBound(reportHours)) As String
        'For i = LBound(reportHours) To UBound(reportHours)
        '    reportHoursStr(i) = format$(reportHours(i), "00")
        'Next i
    


        ' Build query params
        Dim queryData As String
        Dim myURL As String
        
        
        
        If techType = "DO" Then
            ' Cell list needs to be formatted as 4 digits
            If IsArray(cellList) Then cellList = Array_ToString(cellList, "0000")
            
            
            queryData = URL_BuildQueryString("market", market, _
                "emon", Format$(dateEnd, "mm"), "eday", Format$(dateEnd, "dd"), "eyr", Format$(dateEnd, "yyyy"), _
                "numdays", Format$(numberOfDaysToTrend, "00"), _
                "hr", reportHoursStr, _
                "dw", reportDays, _
                "cell", cellList, _
                "s", reportSectors, _
                "carr", "all", _
                "type", reportType, _
                "gby", reportGroupBy, _
                "con", reportContent, _
                "ordby", "datehr", "orddir", "asc", _
                 "ntype", "eq", "n", "0", "threshold", "none", "throp", "gt", "thvalue", "0", _
                "bsc", "-All-", _
                 "maxrows", "128000", "newwin", "on" _
            )
        
            myURL = BASE_URI & "/cgi-bin/evdo.cgi?" & queryData
        ElseIf techType = "1X" Then
            Dim cbscList As Variant
            
            cbscList = Array(14, 16, 18, 31, 32, 33, 34, 35, 36, 37) 'Central PA
        
            queryData = URL_BuildQueryString("market", market, _
                "emon", Format$(dateEnd, "mm"), "eday", Format$(dateEnd, "dd"), "eyr", Format$(dateEnd, "yyyy"), _
                "numdays", Format$(numberOfDaysToTrend, "00"), _
                "hr", reportHoursStr, _
                "dw", reportDays, _
                "bts", cellList, _
                "s", reportSectors, _
                "c", "all", _
                "type", reportType, _
                "gby", reportGroupBy, _
                "con", reportContent, _
                "cbsc", cbscList, _
                "numcbs", Array_Count(cbscList), _
                "ordby", "datehr", "orddir", "asc", "threshold", "none", "throp", "gt", "thvalue", "0", _
                "ntype", "eq", "n", "0", _
                "maxrows", "100000", "newwin", "on" _
            )
        
            myURL = BASE_URI & "/cgi-bin/sectcdl.cgi?" & queryData
        End If
        
        

        Debug.Print "MPT " & techType & ": " & myURL
    

        Dim httpReq As New WinHttp.WinHttpRequest
        
        
        httpReq.SetTimeouts 0, -1, -1, -1 '0, 60000, 30000, 60000 ' Unlimited, 60s (default), 30s (default), 60s (2 x default)
        httpReq.Open "GET", myURL, False
        httpReq.Send
        
        
        
    
        ' TODO: Check result InStr Your query returned no data.
        
    
    
        Dim dataExcelPath As String
        dataExcelPath = ExtractStringBetweenText(httpReq.ResponseText, "<a href=", ">Excel Import</a>", True)
        
        Set httpReq = Nothing
        
        
        
        myURL = BASE_URI & dataExcelPath
        
        Debug.Print myURL
        
    
        GetWebData myURL, destWorksheet
        
    
    

End Function
Public Sub Cluster_CellGroupReport_1XDO_Test()
' Test function - Do not use

    Dim clusterDef_1XDO As String
    
    clusterDef_1XDO = "8-4-998-2,3; 4-220-1; 56-4-2"
    
    Dim mtsoRefTable() As MTSO_Reference_Table_Item
    mtsoRefTable = MTSO_Reference_Table_BuildFromRange(ThisWorkbook.Names("MTSO_TABLE").RefersToRange)


    Cluster_CellGroupReport_1XDO clusterDef_1XDO:=clusterDef_1XDO, _
        dateEnd:=Now() - 1, daysToTrend:=2, hours:=Array(6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23), _
            mtsoRefTable:=mtsoRefTable, _
            NCWS_1X_worksheetName:="NCWS_1X_T", _
            NCWS_1X_CELL_worksheetName:="NCWS_1X_CELL_T", _
            NCWS_DO_worksheetName:="NCWS_DO_T", _
            MPT_1X_worksheetName:="MPT_1X_T", MPT_1X_reportContent:=Array("perf", "corp"), _
            MPT_DO_worksheetName:="MPT_DO_T", MPT_DO_reportContent:=Array("perf", "evdothruput", "tch_usage", "blocking")

End Sub

' Cluster_CellGroupReport_1XDO: Runs multiple 1X/DO reports given a 1X/DO cluster. Will use either NCWS or MPT reporting tool depending on the market and mtsoRefTable()
'   The expected format for clusterDef_1XDO is ECP-CELL-SECTOR1,SECTOR2; ECP-CELL-SECTOR; ....
Public Sub Cluster_CellGroupReport_1XDO(clusterDef_1XDO As Variant, dateEnd As Date, daysToTrend As Integer, hours As Variant, mtsoRefTable() As MTSO_Reference_Table_Item, _
    Optional NCWS_1X_worksheetName = "NCWS_1X", Optional NCWS_1X_reportType = "hourly", Optional NCWS_1X_reportContent As Variant = 1, _
    Optional NCWS_1X_CELL_worksheetName = "NCWS_1X_CELL", Optional NCWS_1X_CELL_reportType = "hourly", Optional NCWS_1X_CELL_reportContent As Variant = 1, _
    Optional NCWS_DO_worksheetName = "NCWS_DO", Optional NCWS_DO_reportType = "hourly", Optional NCWS_DO_reportContent As Variant = 1, _
    Optional MPT_1X_worksheetName = "MPT_1X", Optional MPT_1X_reportType = "hourly", Optional MPT_1X_reportContent As Variant = "perf", _
    Optional MPT_DO_worksheetName = "MPT_DO", Optional MPT_DO_reportType = "hourly", Optional MPT_DO_reportContent As Variant = "perf" _
    )
    
   
    
    
    Dim i As Long, j As Long
    Dim clusterDefList_1XDO() As String
    
    
    
    If IsArray(clusterDef_1XDO) Then
        If Array_Count(clusterDef_1XDO) = 0 Then Err.Raise -1, , "Cluster definition (1X/DO) is a required parameter"
        
        clusterDefList_1XDO = Array_ToString(clusterDef_1XDO)
    Else
        If IsMissing(clusterDef_1XDO) Or clusterDef_1XDO = "" Then
            Err.Raise -1, , "Cluster definition (1X/DO) is a required parameter"
        Else
            clusterDefList_1XDO = ClusterDef_Expand(CStr(clusterDef_1XDO))
        End If
    End If
    
    

    
    ' -----------------------------------------------------------------------------------------------
    ' Begin - Convert 1X/DO clusters in their correct form for their respective reporting tool and tech type, grouping each by market
    ' -----------------------------------------------------------------------------------------------
    Dim cluster_NCWS_1X_ByMarket As Variant, cluster_NCWS_1X_Markets As Variant, cluster_NCWS_1X_MarketECPs As Variant
    Dim cluster_NCWS_DO_ByMarket As Variant, cluster_NCWS_DO_Markets As Variant, cluster_NCWS_DO_MarketECPs As Variant
    Dim cluster_MPT_1X_ByMarket As Variant, cluster_MPT_1X_Markets As Variant, cluster_MPT_1X_MarketECPs As Variant
    Dim cluster_MPT_DO_ByMarket As Variant, cluster_MPT_DO_Markets As Variant, cluster_MPT_DO_MarketECPs As Variant
    
    Dim clusterItem As Variant, clusterItemParts() As String
    Dim clusterItem_SysID As Integer, clusterItem_ECP As Integer, clusterItem_SN As Integer, clusterItem_Cell, clusterItem_Sector As Integer
    
    
    
    
    Dim marketRef As MTSO_Reference_Table_Item
    
    Dim clusterListItem As String
    Dim cellListItem As String
    Dim foundInList As Boolean
    Dim marketIdx As Long
    
    Dim mtsoIdx As Long

    
    For i = LBound(clusterDefList_1XDO) To UBound(clusterDefList_1XDO)
        clusterItemParts = Split(clusterDefList_1XDO(i), "-") ' Assume format: [SysID]-<ECP>-<Cell>-<Sector>
        'clusterItem_ECP = CInt(clusterItemParts(0))
        'clusterItem_Cell = CInt(clusterItemParts(1))
        'clusterItem_Sector = CInt(clusterItemParts(2))
        clusterItem_Sector = CInt(clusterItemParts(UBound(clusterItemParts) - 0))
        clusterItem_Cell = CInt(clusterItemParts(UBound(clusterItemParts) - 1))
        clusterItem_ECP = CInt(clusterItemParts(UBound(clusterItemParts) - 2))
        
        clusterItem_SysID = 0
        If Array_Count(clusterItemParts) > 3 Then clusterItem_SysID = CInt(clusterItemParts(UBound(clusterItemParts) - 3))
        
        ' find mtso ref table item by ECP
        mtsoIdx = -1
        For j = LBound(mtsoRefTable) To UBound(mtsoRefTable)
            If mtsoRefTable(j).ECP = clusterItem_ECP And (clusterItem_SysID = mtsoRefTable(j).SysID Or clusterItem_SysID = 0) Then mtsoIdx = j
        Next j
        ' if mtso ref table item is not found by ECP, then perhaps the cluster format is [SysID]-<SN>-<Cell>-<Sector> (e.g. the ECP is actually the SN)
        If mtsoIdx < 0 Then
            For j = LBound(mtsoRefTable) To UBound(mtsoRefTable)
                If mtsoRefTable(j).SN = clusterItem_ECP And (clusterItem_SysID = mtsoRefTable(j).SysID Or clusterItem_SysID = 0) Then mtsoIdx = j
            Next j
        End If
        
        
        If mtsoIdx < 0 Then Err.Raise -1, , "Market with SysID and ECP/SN '" & IIf(clusterItem_SysID > 0, clusterItem_SysID & "-", "") & clusterItem_ECP & "' not found in reference table"
            
        marketRef = mtsoRefTable(mtsoIdx)
        
        
        
        If marketRef.ReportingTool_1XDO = "MPT" Then

            ' Begin - Create MPT 1X Cluster
            clusterListItem = clusterItem_Cell & "-" & clusterItem_Sector
            
    
            ' Find and add market if doesn't exist
            marketIdx = Array_Find(cluster_MPT_1X_Markets, marketRef.MPT_1XDO_MarketName)
            If marketIdx = -1 Then
                cluster_MPT_1X_Markets = Array_Append(cluster_MPT_1X_Markets, marketRef.MPT_1XDO_MarketName)
                cluster_MPT_1X_MarketECPs = Array_Append(cluster_MPT_1X_MarketECPs, marketRef.ECP)
                
                ' Now add cluster item to market
                cluster_MPT_1X_ByMarket = Array_Append(cluster_MPT_1X_ByMarket, Array(clusterListItem))
            Else
                cluster_MPT_1X_ByMarket(marketIdx) = Array_Append(cluster_MPT_1X_ByMarket(marketIdx), clusterListItem)
            End If
            ' End - Create MPT 1X Cluster
            
            
            
            ' Begin - Create MPT DO Cluster
            clusterListItem = clusterItem_Cell & "-" & clusterItem_Sector
            
    
            ' Find and add market if doesn't exist
            marketIdx = Array_Find(cluster_MPT_DO_Markets, marketRef.MPT_1XDO_MarketName)
            If marketIdx = -1 Then
                cluster_MPT_DO_Markets = Array_Append(cluster_MPT_DO_Markets, marketRef.MPT_1XDO_MarketName)
                cluster_MPT_DO_MarketECPs = Array_Append(cluster_MPT_DO_MarketECPs, marketRef.ECP)
                
                ' Now add cluster item to market
                cluster_MPT_DO_ByMarket = Array_Append(cluster_MPT_DO_ByMarket, Array(clusterListItem))
            Else
                cluster_MPT_DO_ByMarket(marketIdx) = Array_Append(cluster_MPT_DO_ByMarket(marketIdx), clusterListItem)
            End If
            ' End - Create MPT DO Cluster
            
            

        ElseIf marketRef.ReportingTool_1XDO = "NCWS" Then
            
                    
            ' Begin - Create NCWS DO Cluster
            clusterListItem = marketRef.SysID & "-" & marketRef.SN & "-" & clusterItem_Cell & "-" & clusterItem_Sector
            
    
            ' Find and add market if doesn't exist
            marketIdx = Array_Find(cluster_NCWS_DO_Markets, marketRef.NCWS_1XDO_CellGroupName)
            If marketIdx = -1 Then
                cluster_NCWS_DO_Markets = Array_Append(cluster_NCWS_DO_Markets, marketRef.NCWS_1XDO_CellGroupName)
                cluster_NCWS_DO_MarketECPs = Array_Append(cluster_NCWS_DO_MarketECPs, marketRef.ECP)
                
                ' Now add cluster item to market
                cluster_NCWS_DO_ByMarket = Array_Append(cluster_NCWS_DO_ByMarket, Array(clusterListItem))
            Else
                cluster_NCWS_DO_ByMarket(marketIdx) = Array_Append(cluster_NCWS_DO_ByMarket(marketIdx), clusterListItem)
            End If
            ' End - Create NCWS DO Cluster
            
            
            ' Begin - Create NCWS 1X Cluster
            clusterListItem = marketRef.SysID & "-" & marketRef.ECP & "-" & clusterItem_Cell & "-" & clusterItem_Sector
            
    
            ' Find and add market if doesn't exist
            marketIdx = Array_Find(cluster_NCWS_1X_Markets, marketRef.NCWS_1XDO_CellGroupName)
            If marketIdx = -1 Then
                cluster_NCWS_1X_Markets = Array_Append(cluster_NCWS_1X_Markets, marketRef.NCWS_1XDO_CellGroupName)
                cluster_NCWS_1X_MarketECPs = Array_Append(cluster_NCWS_1X_MarketECPs, marketRef.ECP)
                
                ' Now add cluster item to market
                cluster_NCWS_1X_ByMarket = Array_Append(cluster_NCWS_1X_ByMarket, Array(clusterListItem))
            Else
                cluster_NCWS_1X_ByMarket(marketIdx) = Array_Append(cluster_NCWS_1X_ByMarket(marketIdx), clusterListItem)
            End If
            ' End - Create NCWS 1X Cluster
            
            
            
        End If
        
        
    Next i
    
    ' -----------------------------------------------------------------------------------------------
    ' End - Convert 1X/DO clusters in their correct form for their respective reporting tool and tech type, grouping each by market
    ' -----------------------------------------------------------------------------------------------
    
    Debug.Print
    
    
    ' -----------------------------------------------------------------------------------------------
    ' Begin - Load
    ' -----------------------------------------------------------------------------------------------
    Dim reportingTool As String ',reportingToolIdx As Integer,
    Dim techType As String ',techTypeIdx As Integer,
    Dim reportType As String, reportGroupBy As String, reportContent As Variant
    
    Dim marketName As String, marketECP As Long
    Dim marketCluster As Variant, cellList As Variant
    Dim destWorksheetName As String, wsDest As Worksheet
    Dim destLastRow As Long
    Dim cluster_ByMarket As Variant, cluster_Markets As Variant, cluster_MarketECPs As Variant
    
    
    ' Loop through each reporting tool (NCWS, MPT) and each technology type (1X,DO). Download the data into their respective worksheets.
    Dim passNo As Integer

    
    For passNo = 1 To 5
        ' Reset
        reportingTool = "": techType = ""
        reportGroupBy = "": reportType = "": reportContent = Null
        cluster_ByMarket = Null
        cluster_Markets = Null
        cluster_MarketECPs = Null
        destWorksheetName = ""
    
    
        ' Passes:
        '    #  Tool    Tech    Level
        '    -------------------------
        '    1: NCWS    1X      Sector
        '    2: NCWS    1X      Cell
        '    3: NCWS    DO      Sector
        '    4: MPT     1X      Sector
        '    5: MPT     DO      Sector
        Select Case passNo
            Case 1:     '   NCWS / 1X / Sector Level
                reportingTool = "NCWS": techType = "1X"
                reportGroupBy = "Sector"
                reportType = NCWS_1X_reportType
                reportContent = NCWS_1X_reportContent
                cluster_ByMarket = cluster_NCWS_1X_ByMarket
                cluster_Markets = cluster_NCWS_1X_Markets
                cluster_MarketECPs = cluster_NCWS_1X_MarketECPs
                destWorksheetName = NCWS_1X_worksheetName
            Case 2:     '   NCWS / 1X / Cell Level
                reportingTool = "NCWS": techType = "1X"
                reportGroupBy = "Cell"
                reportType = NCWS_1X_CELL_reportType
                reportContent = NCWS_1X_CELL_reportContent
                cluster_ByMarket = cluster_NCWS_1X_ByMarket
                cluster_Markets = cluster_NCWS_1X_Markets
                cluster_MarketECPs = cluster_NCWS_1X_MarketECPs
                destWorksheetName = NCWS_1X_CELL_worksheetName
            Case 3:     '   NCWS / DO / Sector Level
                reportingTool = "NCWS": techType = "DO"
                reportGroupBy = "Sector"
                reportType = NCWS_DO_reportType
                reportContent = NCWS_DO_reportContent
                cluster_ByMarket = cluster_NCWS_DO_ByMarket
                cluster_Markets = cluster_NCWS_DO_Markets
                cluster_MarketECPs = cluster_NCWS_DO_MarketECPs
                destWorksheetName = NCWS_DO_worksheetName
            Case 4:     '   MPT / 1X / Sector Level
                reportingTool = "MPT": techType = "1X"
                reportGroupBy = "sector"
                reportType = MPT_1X_reportType
                reportContent = MPT_1X_reportContent
                cluster_ByMarket = cluster_MPT_1X_ByMarket
                cluster_Markets = cluster_MPT_1X_Markets
                cluster_MarketECPs = cluster_MPT_1X_MarketECPs
                destWorksheetName = MPT_1X_worksheetName
            Case 5:     '   MPT / DO / Sector Level
                reportingTool = "MPT": techType = "DO"
                reportGroupBy = "sector"
                reportType = MPT_DO_reportType
                reportContent = MPT_DO_reportContent
                cluster_ByMarket = cluster_MPT_DO_ByMarket
                cluster_Markets = cluster_MPT_DO_Markets
                cluster_MarketECPs = cluster_MPT_DO_MarketECPs
                destWorksheetName = MPT_DO_worksheetName
        End Select

        'reportingTool = Choose(passNo, "NCWS", "NCWS", "NCWS", "MPT", "MPT")
        'techType = Choose(passNo, "1X", "1X", "DO", "1X", "DO")
        'reportGroupBy = Choose(passNo, "Sector", "Cell", "Sector", "sector", "sector")
        
        If destWorksheetName <> "" Then
            ' Sanity check
            Debug.Assert reportingTool <> "":
            Debug.Assert techType <> ""
            Debug.Assert reportGroupBy <> ""
            Debug.Assert reportType <> ""
            Debug.Assert IsNull(reportContent) = False
            Debug.Assert IsNull(cluster_ByMarket) = False
            Debug.Assert IsNull(cluster_Markets) = False
            Debug.Assert IsNull(cluster_MarketECPs) = False
            'Debug.Assert destWorksheetName <> ""
        
        
    
            If Array_Count(cluster_ByMarket) > 0 Then
                
                Application.ScreenUpdating = False
                Application.Calculation = xlCalculationManual
                
                
                If Not Worksheet_SheetExists(destWorksheetName) Then
                    Set wsDest = ThisWorkbook.Sheets.Add
                    wsDest.Name = destWorksheetName
                    destLastRow = 0
                Else
                    Set wsDest = ThisWorkbook.Sheets(destWorksheetName)
                    wsDest.Cells.ClearContents
                    destLastRow = Worksheet_GetLastRow(wsDest)
                End If
                
                
                For i = LBound(cluster_ByMarket) To UBound(cluster_ByMarket)
                   
                    marketName = cluster_Markets(i)
                    marketECP = cluster_MarketECPs(i)
                    marketCluster = cluster_ByMarket(i)
                   
                
                    Dim tmpWorksheetName As String, tmpWS As Worksheet
                    
                    tmpWorksheetName = "tempData" & Format(Now, "yyyymmddhhmmss")
                    
                    If reportingTool = "NCWS" Then
                    
                
                        ' NCWS
                        If UCase(reportGroupBy) = "CELL" Then
                           cellList = ClusterDef_Extract_UniqueCellList(marketCluster)
                           
                           NCWS_CellGroupReport destWorksheet:=tmpWorksheetName, techType:=techType, _
                               cellGroup:=marketName, cellList:=cellList, _
                               reportType:=reportType, reportGroupBy:=reportGroupBy, reportContent:=reportContent, _
                               dateEnd:=dateEnd, numberOfDaysToTrend:=daysToTrend, reportHours:=hours
                        Else
                           NCWS_CellGroupReport_Cluster destWorksheet:=tmpWorksheetName, techType:=techType, _
                               cellGroup:=marketName, clusterDef:=marketCluster, _
                               reportType:=reportType, reportGroupBy:=reportGroupBy, reportContent:=reportContent, _
                               dateEnd:=dateEnd, numberOfDaysToTrend:=daysToTrend, reportHours:=hours
                       End If
                            
                    ElseIf reportingTool = "MPT" Then
                        
                        ' MPT
                        MPT_CellGroupReport_Cluster destWorksheet:=tmpWorksheetName, techType:=techType, _
                            market:=marketName, clusterDef:=marketCluster, _
                            reportType:=reportType, reportGroupBy:=reportGroupBy, reportContent:=reportContent, _
                            dateEnd:=dateEnd, numberOfDaysToTrend:=daysToTrend, reportHours:=hours
                            
                    End If
                    
                    
                
                    Set tmpWS = ThisWorkbook.Sheets(tmpWorksheetName)
                
                    Dim tmpLastRow As Long, tmpLastCol As Long
                    Dim colNbr_BTS As Long
                    Dim rngECP As Range
                    
                    
                    tmpLastRow = Worksheet_GetLastRow(tmpWS)
                    tmpLastCol = Worksheet_GetLastColumn(tmpWS)
                    
                    
                    ' Begin - Create ECP column and fill with ECP (MPT/1X only)
                    If reportingTool = "MPT" And techType = "1X" Then
                        colNbr_BTS = Worksheet_GetColumnByName(tmpWS, "BTS")
                        Debug.Assert colNbr_BTS > 0
                        
                        tmpWS.columns(colNbr_BTS).Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
                        
                        Set rngECP = tmpWS.columns(colNbr_BTS)
                        
                        
                        rngECP.Cells(1, 1) = "ECP"
                        Set rngECP = rngECP.Cells(1, 1).Offset(1, 0).Resize(tmpLastRow - 1, 1)
                        
                        rngECP.value = Array_CreateAndFill(tmpLastRow - 1, marketECP)
                        
                        tmpLastCol = tmpLastCol + 1
                    End If
                    ' End - Create ECP column and fill with ECP (MPT/1X only)
                    
                    
                    ' Begin - Append data in temporary sheet to final worksheet
                    Dim destRng As Range
                    Dim srcRng As Range
                    
                        
                    If destLastRow = 0 Then
                        ' First row - include header columns
                        Set srcRng = tmpWS.Cells(1, 1).Resize(tmpLastRow, tmpLastCol)
                        Set destRng = wsDest.Cells(1, 1).Resize(tmpLastRow, tmpLastCol)
                    Else
                        Set srcRng = tmpWS.Cells(1, 1).Offset(1).Resize(tmpLastRow - 1, tmpLastCol)
                        Set destRng = wsDest.Cells(1, 1).Offset(destLastRow).Resize(tmpLastRow - 1, tmpLastCol)
                    End If
                
                    Dim data() As Variant
                    
                    data = srcRng
                    destRng = data
                    
                    ' End - Append data in temporary sheet to final worksheet
                    
                    Application.DisplayAlerts = False
                    tmpWS.Delete
                    Application.DisplayAlerts = True
                        
                    destLastRow = destLastRow + tmpLastRow - 1
                    
                    Debug.Print
                Next i
                
                
                Application.ScreenUpdating = True
                Application.Calculation = xlCalculationAutomatic
                
    
            End If ' Array_Count(cluster_ByMarket) > 0

        End If  ' destWorksheetName <> "" ?

        
    Next passNo
 

End Sub
' Creates two arrays which represent valid date ranges for an event, given the event start/end dates and hours. Needed because hours can sometimes be specified past midnight
'
Public Function Event_CreateDateRanges(eventStartDate As Date, eventEndDate As Date, eventDailyStartHour As Integer, eventDailyEndHour As Integer, eventDaysOfWeek As Variant) As Date()


    'Dim eventStartDate As Date, eventEndDate As Date, eventDailyStartHour As Integer, eventDailyEndHour As Integer, eventDaysOfWeek As Variant
    'eventStartDate = #10/4/2014#
    'eventEndDate = #10/5/2014#
    'eventDailyStartHour = 10
    'eventDailyEndHour = 16 ' 5PM
    'eventDaysOfWeek = Array(VbDayOfWeek.vbSaturday, VbDayOfWeek.vbSunday)
    
    If IsArray(eventDaysOfWeek) = False Then eventDaysOfWeek = Array(1, 2, 3, 4, 5, 6, 7)



    Dim validWeekDay As Variant
    Dim i As Long
    
    Dim curDate As Date
    Dim hourDiff As Integer
    Dim isValidDayOfWeek As Boolean
    
    Dim dateRanges() As Date
    Dim dateRangesCount As Integer
    Dim dateRangeStart As Date, dateRangeEnd As Date
    
    hourDiff = eventDailyEndHour - eventDailyStartHour
     ' Check if event end time is into the next day (e.g. 7pm-1am)
    If hourDiff < 0 Then hourDiff = hourDiff + 24
       
        
    
    
    dateRangesCount = 0
    curDate = eventStartDate
    
    Do While curDate <= eventEndDate
    

        isValidDayOfWeek = False
        
        For i = LBound(eventDaysOfWeek) To UBound(eventDaysOfWeek)
            validWeekDay = eventDaysOfWeek(i)
            
            If Weekday(curDate) = validWeekDay Then isValidDayOfWeek = True: Exit For
        Next i
        
        If isValidDayOfWeek Then
            dateRangesCount = dateRangesCount + 1
           
            ReDim Preserve dateRanges(1 To 2, 1 To dateRangesCount) As Date
            
            dateRangeStart = DateAdd("h", eventDailyStartHour, DateSerial(Year(curDate), Month(curDate), day(curDate)))
            dateRangeEnd = DateAdd("h", hourDiff, dateRangeStart)
            
            dateRanges(1, dateRangesCount) = dateRangeStart
            dateRanges(2, dateRangesCount) = dateRangeEnd
        End If
    
        curDate = DateAdd("d", 1, curDate)
    Loop
    
    
    ' Transpose dateRanges array
    Dim retArr() As Date
    ReDim retArr(1 To dateRangesCount, 1 To 2) As Date
    
    For i = 1 To dateRangesCount
        retArr(i, 1) = dateRanges(1, i)
        retArr(i, 2) = dateRanges(2, i)
    Next i
    

    Event_CreateDateRanges = retArr
    

End Function


' Converts string like "Sun/Mon/Tue/Wed" to an array of integers (1=Sun, 2=Mon, 3=Tue,...)
Public Function Event_ParseDaysOfWeek(daysOfWeekStr As String, Optional default As Variant) As Variant

    'Dim daysOfWeekStr As String: daysOfWeekStr = "Su/Mo/Tue/Wed"
    'Dim default As Variant
    
    If IsMissing(default) Then default = Array()
    
    If Len(daysOfWeekStr) = 0 Then
        Event_ParseDaysOfWeek = default
        Exit Function
    End If
    
    
    Dim arrSplit As Variant, day As Variant
    Dim daysOfWeekBitfield As Byte
    Dim dayCount As Integer
    Dim i As Integer
    
    
    daysOfWeekBitfield = 0
    dayCount = 0
    
    'Split by comma
    arrSplit = Split(UCase(daysOfWeekStr), ",")
    

    
    If Array_Count(arrSplit) = 1 Then arrSplit = Split(UCase(daysOfWeekStr), "/")
    
    ' Use bitfield to store days of week
    For Each day In arrSplit
        If day Like "SU*" Then daysOfWeekBitfield = daysOfWeekBitfield Or 1
        If day Like "M*" Then daysOfWeekBitfield = daysOfWeekBitfield Or 2
        If day Like "TU*" Then daysOfWeekBitfield = daysOfWeekBitfield Or 4
        If day Like "W*" Then daysOfWeekBitfield = daysOfWeekBitfield Or 8
        If day Like "TH*" Then daysOfWeekBitfield = daysOfWeekBitfield Or 16
        If day Like "F*" Then daysOfWeekBitfield = daysOfWeekBitfield Or 32
        If day Like "SA*" Then daysOfWeekBitfield = daysOfWeekBitfield Or 64
    Next day
    
    For i = 0 To 6: dayCount = dayCount + IIf(daysOfWeekBitfield And (2 ^ i), 1, 0): Next i
    
    If dayCount = 0 Then
        Event_ParseDaysOfWeek = default
        Exit Function
    End If
    
    Dim daysOfWeekArr As Variant, dayIdx As Integer
    
    ReDim daysOfWeekArr(1 To dayCount) As Integer
    
    dayIdx = 1
    
    For i = 0 To 6
        If daysOfWeekBitfield And (2 ^ i) Then
            daysOfWeekArr(dayIdx) = i + 1
            dayIdx = dayIdx + 1
        End If
    Next i
    
    Event_ParseDaysOfWeek = daysOfWeekArr
    
End Function

Public Sub Event_FilterWorksheetByDateRange(ws As Worksheet, colDate As Variant, colHr As Variant, dateRanges() As Date)

    ' 2015-05-13 - Allow columns to be specified by name

    Debug.Assert LBound(dateRanges, 2) = 1
    Debug.Assert UBound(dateRanges, 2) = 2

    
    ' No point in moving forward if there is no data or only a header row
    If Worksheet_GetLastRow(ws) <= 1 Then Exit Sub
    
    
    Dim colNbr_Date As Long, colNbr_Hr As Long
    
    If IsNumeric(colDate) Then
        colNbr_Date = colDate
    Else
        colNbr_Date = Worksheet_GetColumnByName(ws, CStr(colDate))
    End If
    
    If IsNumeric(colHr) Then
        colNbr_Hr = colHr
    Else
        colNbr_Hr = Worksheet_GetColumnByName(ws, CStr(colHr))
    End If
    
    Debug.Assert colNbr_Date > 0
    Debug.Assert colNbr_Hr > 0
    
    
    Dim i As Integer, j As Integer
    Dim colNbr_Last As Long, rowNbr_Last As Long
    
    
    colNbr_Last = Worksheet_GetLastColumn(ws)
    rowNbr_Last = Worksheet_GetLastRow(ws)
    
    
    Application.ScreenUpdating = False ' Stop screen painting (speeds up processing)
    
    Dim isDateWithinAnyRange As Boolean
    
    For i = rowNbr_Last To 2 Step -1
        Dim row_DateTime As Date
        
        row_DateTime = DateValue(ws.Cells(i, colNbr_Date))
        row_DateTime = DateAdd("h", ws.Cells(i, colNbr_Hr), row_DateTime) 'Assumes hour is 0-23
        

        isDateWithinAnyRange = False
        
        For j = LBound(dateRanges, 1) To UBound(dateRanges, 1)
            ' dateRanges(j, 1) = Start Date/Time, dateRanges(j, 2) = End Date/Time
            If dateRanges(j, 1) <= row_DateTime And row_DateTime <= dateRanges(j, 2) Then
                isDateWithinAnyRange = True
                Exit For
            End If
        Next j
        
        
    
        If isDateWithinAnyRange = False Then
            ws.Rows(i).Delete
            'ws.Rows(i).Interior.Color = RGB(255, 0, 0)
        End If
    Next i
    
    
    
    
    Application.ScreenUpdating = True ' Repaint screen
    

End Sub

' Utility function to quickly populate items in the MTSO lookup table using a table range
' The first row is the column headings and should contain the following columns:
'  LTE Mkt, Sys ID, SN, ECP, 1X/DO Reporting Tool, 1X/DO CellGroup
Public Function MTSO_Reference_Table_BuildFromRange(tableRng As Range) As MTSO_Reference_Table_Item()



    Dim col_LTE_Mkt As Long, col_SysID As Long, col_SN As Long, col_ECP As Long, col_1XDO_ReportingTool As Long, col_1XDO_CellGroup_MarketName As Long
    
    
    Dim numCols As Long, numRows As Long
    Dim i As Long
    Dim colName As String
    
    numCols = Range_LastColumn(tableRng)
    numRows = Range_LastRow(tableRng)
    
    For i = 1 To numCols
        colName = UCase(CStr(tableRng.Cells(1, i).value))
        
        If colName Like "LTE?MKT*" Or colName Like "LTE?MARKET*" Then col_LTE_Mkt = i
        If colName Like "SYS?ID" Or colName Like "SYSTEM?ID" Then col_SysID = i
        If colName = "SN" Or colName Like "SERVICE?NODE" Then col_SN = i
        If colName = "ECP" Then col_ECP = i
        
        If colName Like "1X*DO*REPORT*TOOL" Then col_1XDO_ReportingTool = i
        
        If colName Like "1X*DO*CELLGROUP*" Or colName Like "1X*DO*MARKET*" Then col_1XDO_CellGroup_MarketName = i
    Next i
    
    
    Dim mtsoRefTable() As MTSO_Reference_Table_Item
    ReDim mtsoRefTable(1 To numRows - 1) As MTSO_Reference_Table_Item

    
    For i = 2 To numRows
        If IsNumeric(tableRng.Cells(i, col_LTE_Mkt)) = False Then Err.Raise -1, , "MTSO Reference Table: Invalid LTE Mkt ID: " & tableRng.Cells(i, col_LTE_Mkt)
        If IsNumeric(tableRng.Cells(i, col_SN)) = False Then Err.Raise -1, , "MTSO Reference Table: Invalid SN: " & tableRng.Cells(i, col_SN)
        If IsNumeric(tableRng.Cells(i, col_ECP)) = False Then Err.Raise -1, , "MTSO Reference Table: Invalid ECP: " & tableRng.Cells(i, col_ECP)

        

        MTSO_Reference_Table_PopulateItem mtsoRefTable:=mtsoRefTable, idx:=i - 1, _
            LTE_Market_ID:=tableRng.Cells(i, col_LTE_Mkt), _
            SysID:=tableRng.Cells(i, col_SysID), _
            SN:=tableRng.Cells(i, col_SN), _
            ECP:=tableRng.Cells(i, col_ECP), _
            rptTool_1XDO:=tableRng.Cells(i, col_1XDO_ReportingTool), _
            rptTool_marketOrCellGroupName_1XDO:=tableRng.Cells(i, col_1XDO_CellGroup_MarketName)
            
    Next i
    
    
    MTSO_Reference_Table_BuildFromRange = mtsoRefTable

End Function
' Utility function to quickly populate items in the MTSO lookup table
Public Sub MTSO_Reference_Table_PopulateItem(mtsoRefTable() As MTSO_Reference_Table_Item, idx As Integer, LTE_Market_ID As Integer, SysID As Integer, SN As Integer, ECP As Integer, rptTool_1XDO As String, rptTool_marketOrCellGroupName_1XDO As String)
    
    
    mtsoRefTable(idx).LTE_Market_ID = LTE_Market_ID
    mtsoRefTable(idx).SysID = SysID
    mtsoRefTable(idx).SN = SN
    mtsoRefTable(idx).ECP = ECP
    
    
    mtsoRefTable(idx).ReportingTool_1XDO = rptTool_1XDO
    
    If rptTool_1XDO = "MPT" Then
        mtsoRefTable(idx).MPT_1XDO_MarketName = rptTool_marketOrCellGroupName_1XDO
    ElseIf rptTool_1XDO = "NCWS" Then
        mtsoRefTable(idx).NCWS_1XDO_CellGroupName = rptTool_marketOrCellGroupName_1XDO
    Else
        Err.Raise -1, , "Invalid reporting tool '" & rptTool_1XDO & "'. Reporting tool is either MPT or NCWS"
    End If
    
End Sub
