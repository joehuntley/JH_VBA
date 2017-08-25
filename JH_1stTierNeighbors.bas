Attribute VB_Name = "JH_1stTierNeighbors"
Option Explicit


' JH_1stTierNeighbors
' ------------------------------------------------------------------------------------------------------------------
' sync 20150707
' ------------------------------------------------------------------------------------------------------------------'
' VBA functions which support the 1st tier neighbor generation algorithm developed to auto-calculate offload sectors
' for a given lat/long.
'
'
' Joseph Huntley (joseph.huntley@vzw.com)
' ------------------------------------------------------------------------------------------------------------------
' Changelog:
'
'
' ------------------------------------------------------------------------------------------------------------------
Private Const EARTH_RADIUS_MI = 3959
Private Const PI As Double = 3.14159265358979
Private Const EPSILON As Double = 0.000001 ' small number for comparing double values

Private Const AZIMUTH_UNDEFINED = -1

Private Const DBG_PRINT_FLAG = False

Type Path_2D
    distance As Double
    angle As Double
    dx As Double
    dy As Double
End Type


Public Type CellSite
    CellNum As Integer
    SwitchID As Integer
    Name As String
    Lat As Double
    Long As Double
    UID As String

    Alpha_Azimuth As Integer
    Beta_Azimuth As Integer
    Gamma_Azimuth As Integer
    
    Alpha_BW As Double
    Beta_BW As Double
    Gamma_BW As Double
    
    'Cellular As CellSite_BandConfig
    'U700 As CellSite_BandConfig
    'PCS_CDMA As CellSite_BandConfig
    'PCS_LTE As CellSite_BandConfig
End Type

Public Type CellSite_BandConfig
    Alpha_Azimuth As Integer
    Beta_Azimuth As Integer
    Gamma_Azimuth As Integer
    
    Alpha_BW As Double
    Beta_BW As Double
    Gamma_BW As Double
End Type

Public Type NeighborCell
    cell As CellSite
    
    distance As Double ' in miles
    
    
    coverage_directivity_magnitude As Double
    coverage_directivity_magnitude_scaled As Double
    coverage_directivity_angle As Double
    coverage_directivity_angle_diff As Double
    
    
    is_offload_alpha As Boolean
    is_offload_beta As Boolean
    is_offload_gamma As Boolean
    
    
    ' cell weighting by distance (<distance to neighbor>/<combined distance of all neighbors>) [Range: 0-100%]
    cell_weight As Double
    
    cell_weight2 As Double
    
    ' cell weighting by angle difference of neighbor and sector azimuth (normalized by total number of offload sectors)
    sector_weight_alpha As Double
    sector_weight_beta As Double
    sector_weight_gamma As Double
End Type

Private Type PossibleNeighborCell
    neighbor_cell As NeighborCell
    path As Path_2D

    is_neighbor As Boolean

    'is_offload_alpha As Boolean
    'is_offload_beta As Boolean
    'is_offload_gamma As Boolean
    
    'sector_weight_alpha As Double
    'sector_weight_beta As Double
    'sector_weight_gamma As Double
End Type
' Loads GeoPlan Custom report with pertinent fields and creates and returns an array of cell site structures
Public Function LoadCellListFromGeoPlanReport(fileName As String, bandClass As String) As CellSite()


    Dim isTechnologyLTE As Boolean
    Dim colName_Antenna_Azimuth As String


    isTechnologyLTE = (bandClass = "Upper 700 MHz" Or bandClass = "AWS")


    Select Case bandClass
        Case "Upper 700 MHz":
            colName_Antenna_Azimuth = "Upper 700 MHz Azimuth (deg)"
        Case "Cellular":
            colName_Antenna_Azimuth = "850 MHz Azimuth (deg)"
        Case "PCS":
            colName_Antenna_Azimuth = IIf(isTechnologyLTE, "1900 MHz LTE Azimuth (deg)", "1900 MHz CDMA Azimuth (deg)")
        Case "AWS":
            colName_Antenna_Azimuth = "2100 MHz Azimuth (deg)"
        Case Else:
            Err.Raise -1, , "Unexpected band class: " & bandClass
    End Select



    Dim hFile As Integer
    Dim fileData As String
    Dim fileLines() As String

    hFile = FreeFile()
    Open fileName For Binary Access Read As hFile
    fileData = Space$(LOF(hFile))
    Get hFile, , fileData
    Close hFile

    fileLines = Split(fileData, vbLf)
    fileData = ""


    Dim line As Long
    Dim columnNames() As String, processedHeaderCol As Boolean

    ' Go through header rows
    line = LBound(fileLines)
    processedHeaderCol = False

    Do While line <= UBound(fileLines)
        If Not (fileLines(line) = "# Column Keys" Or fileLines(line) = "# Worksheet Data" Or fileLines(line) = "") Then
            If processedHeaderCol = False Then
                columnNames = Split(fileLines(line), vbTab)
                processedHeaderCol = True
            Else
                Exit Do
            End If
        End If

        line = line + 1
    Loop





    Dim colNbr_Switch As Integer
    Dim colNbr_Cell As Integer
    Dim colNbr_Cell_Name As Integer
    Dim colNbr_eNodeB_ID As Integer
    Dim colNbr_Sector As Integer
    Dim colNbr_Lat As Integer
    Dim colNbr_Lon As Integer
    Dim colNbr_Transmission_Type As Integer
    Dim colNbr_Band_Class As Integer
    Dim colNbr_Antenna_Azimuth As Integer


    colNbr_Switch = Array_Find(columnNames, "Switch Number 2")
    colNbr_Cell = Array_Find(columnNames, "Cell Number")
    colNbr_Cell_Name = Array_Find(columnNames, "Cell Name")
    colNbr_eNodeB_ID = Array_Find(columnNames, "eNodeB ID")
    colNbr_Sector = Array_Find(columnNames, "Sector")
    colNbr_Lat = Array_Find(columnNames, "Latitude Degrees (NAD83)")
    colNbr_Lon = Array_Find(columnNames, "Longitude Degrees (NAD83)")
    colNbr_Transmission_Type = Array_Find(columnNames, "Transmission Type")
    colNbr_Band_Class = Array_Find(columnNames, "Band Class")
    colNbr_Antenna_Azimuth = Array_Find(columnNames, colName_Antenna_Azimuth)

    If colNbr_Switch < 0 Then Err.Raise -1, , "Cannot find column: Switch 2"
    If colNbr_Switch < 0 Then Err.Raise -1, , ("Cannot find column: Switch Number 2")
    If colNbr_Cell < 0 Then Err.Raise -1, , ("Cannot find column: Cell Number")
    If colNbr_Cell_Name < 0 Then Err.Raise -1, , ("Cannot find column: Cell Name")
    If colNbr_eNodeB_ID < 0 Then Err.Raise -1, , ("Cannot find column: eNodeB ID")
    If colNbr_Sector < 0 Then Err.Raise -1, , ("Cannot find column: Sector")
    If colNbr_Lat < 0 Then Err.Raise -1, , ("Cannot find column: Latitude Degrees (NAD83)")
    If colNbr_Lon < 0 Then Err.Raise -1, , ("Cannot find column: Longitude Degrees (NAD83)")
    If colNbr_Transmission_Type < 0 Then Err.Raise -1, , ("Cannot find column: Transmission Type")
    If colNbr_Band_Class < 0 Then Err.Raise -1, , ("Cannot find column: Band Class")
    If colNbr_Antenna_Azimuth < 0 Then Err.Raise -1, , ("Cannot find column: " & colName_Antenna_Azimuth)


    Dim cellList() As CellSite
    Dim cellCount As Long
    Dim cellIdx As Long

    cellCount = 0


    Dim i As Long, j As Long
    Dim row() As String

    Dim rowSwitchNum As Integer, rowCellNum As Integer, rowCellName As String
    Dim transmissionType As String
    Dim sectorStr As String, rowBandClass As String
    Dim sectorAzimuth As Integer
    Dim rowLat As Double, rowLon As Double
    Dim cell_uid As String
    Dim enodeID As Long


    Do While line <= UBound(fileLines) 'And line <= 10
        If Not (fileLines(line) = "") Then
            row = Split(fileLines(line), vbTab)

            rowSwitchNum = CInt(row(colNbr_Switch))
            rowCellNum = CInt(row(colNbr_Cell))
            rowCellName = row(colNbr_Cell_Name)
            sectorStr = row(colNbr_Sector)
            transmissionType = row(colNbr_Transmission_Type)
            sectorStr = row(colNbr_Sector)
            rowBandClass = row(colNbr_Band_Class)
            rowLat = CDbl(row(colNbr_Lat))
            rowLon = CDbl(row(colNbr_Lon))

            sectorAzimuth = AZIMUTH_UNDEFINED
            If IsNumeric(row(colNbr_Antenna_Azimuth)) Then sectorAzimuth = CInt(row(colNbr_Antenna_Azimuth))

            enodeID = 0
            If IsNumeric(row(colNbr_eNodeB_ID)) Then enodeID = CLng(row(colNbr_eNodeB_ID))




            If transmissionType = "Normal Sector" And rowBandClass = bandClass Then

                If isTechnologyLTE Then
                    cell_uid = CStr(enodeID)
                Else
                    cell_uid = rowSwitchNum & "-" & rowCellNum
                End If

                ' Is cell already in list?
                cellIdx = 0

                For j = 1 To cellCount
                    If cellList(j).UID = cell_uid Then
                        cellIdx = j
                        Exit For
                    End If
                Next j

                ' Cell not in list -> add
                If cellIdx = 0 Then
                    cellCount = cellCount + 1
                    ReDim Preserve cellList(1 To cellCount) As CellSite

                    cellIdx = cellCount
                    cellList(cellIdx).SwitchID = rowSwitchNum
                    cellList(cellIdx).Name = rowCellName
                    cellList(cellIdx).CellNum = rowCellNum
                    'cellList(cellIdx).ENB_ID = enodeID
                    cellList(cellIdx).UID = cell_uid
                    cellList(cellIdx).Lat = rowLat
                    cellList(cellIdx).Long = rowLon

                    cellList(cellIdx).Alpha_Azimuth = AZIMUTH_UNDEFINED
                    cellList(cellIdx).Beta_Azimuth = AZIMUTH_UNDEFINED
                    cellList(cellIdx).Gamma_Azimuth = AZIMUTH_UNDEFINED
                End If

                If sectorStr = "D1" And cellList(cellIdx).Alpha_Azimuth = AZIMUTH_UNDEFINED And sectorAzimuth <> AZIMUTH_UNDEFINED Then cellList(cellIdx).Alpha_Azimuth = sectorAzimuth
                If sectorStr = "D2" And cellList(cellIdx).Beta_Azimuth = AZIMUTH_UNDEFINED And sectorAzimuth <> AZIMUTH_UNDEFINED Then cellList(cellIdx).Beta_Azimuth = sectorAzimuth
                If sectorStr = "D3" And cellList(cellIdx).Gamma_Azimuth = AZIMUTH_UNDEFINED And sectorAzimuth <> AZIMUTH_UNDEFINED Then cellList(cellIdx).Gamma_Azimuth = sectorAzimuth


            End If



        End If

        line = line + 1
    Loop


    LoadCellListFromGeoPlanReport = cellList


End Function

' Calc1stTierNeighbors
' ---------------------------------------------------------------------------------------------------------------------------------------
' Calculates the 1st tier neighbors cells and sectors for a given latitude and longitude (subject) using a cell list extracted from a geoplan or
' similiar database.
'
' Assumptions:
'    - Two cells sharing the same latitude and longitude are considered co-located (eg. sector split). Both are considered a neighbor if one of them
'      is a neighbor and their sectors are analyzed as if a single cell contained all the sectors.
'    - A cell cannot be considered a neighbor if the distance between the subject and the cell is greater than maxNeighborDistance
'
' Algorithm Overview:
'   A 1st tier neighbor for the subject location is usually the closest cell for a given bearing. The way a human would typically determine the
'   1st tier neighbors of a subject location would be to:
'       (1) plot the subject location and cells
'       (2) moving in a circle, look at the cells which are closest to the subject. The closest cells for a given bearing are the neighbors
'       (3) for each neighbor cells, the 1st tier neighbor sectors are the sectors which are pointing in the direction of the subject
'
'   The algorithm follows a similiar approach, only  it starts by looking at the closest cells first. The closest cell for a given direction
'   is considered a neighbor. That is, if two or more cells have a similiar bearing, only the closest cell is chosen as the 1st tier neighbor and
'   the rest are disqualified as being a neighbor, provided the latter cells are sufficiently further than the closest cell ie. their distance
'   difference is greater than SAME_BEARING_MIN_SEPARATION_ANGLE_DIST
'
' Algorithm Steps (Determine Neighbor Cells):
'  (1) Create list of viable neighbors - cells which are not co-located but within a given radius specified by maxNeighborDistance
'  (2) Sort viable neighbors by their distance from the subject location. Compare each viable neighbor (in order of distance) to neighbors which are closer
'
'   - TO BE CONTINUED
'
' ---------------------------------------------------------------------------------------------------------------------------------------
Public Function Calc1stTierNeighbors(cellList() As CellSite, sLat As Double, sLong As Double, Optional maxNeighborDistance As Double = 5, Optional includeOnlyNeighborsWithOffloadSectors As Boolean = True) As NeighborCell()


    Const CELL_MIN_DISTANCE = 0.001      ' in miles, minimum distance between cells to be considered a separate cell
    'Const CELL_MAX_DISTANCE = 5          ' in miles, maximum distance between cells to be considered a neighbor
    
    Const CELL_EFFECTIVE_RADIUS = 1.25   ' in miles, - a highly conservative estimate of a cell's effective radius
    
    ' in miles, the minimum radial distance from solution location between two neighbor cells for one of the cells to disqualify the other as a neighbor
    Const SAME_BEARING_MIN_SEPARATION_ANGLE_DIST = 0.25


    ' DEBUG STUFF
    'Dim c  As Variant: c = Array(39.73875, -75.4541)

    ' North Percy: 39.962208, -75.153278
    ' Poplar St: 39.965778, -75.132778
    ' Monticello: 39.937361, -77.687464
    ' WIL PERSES: 39.73875, -75.4541


    'Dim cellList() As CellSite: cellList = LoadCellListFromGeoPlanReport()
    'Dim sLat As Double, sLong As Double
    'sLat = c(0)
    'sLong = c(1)

    'Dim includeOnlyNeighborsWithOffloadSectors As Boolean: includeOnlyNeighborsWithOffloadSectors = True



    Debug.Assert UBound(cellList) > -1


    Dim cellCount As Long

    cellCount = UBound(cellList) - LBound(cellList) + 1


    


    Dim i As Long, j As Long, k As Long
    Dim nIdx As Long, pIdx As Long


    Dim path As Path_2D

    Dim viableNeighbors() As PossibleNeighborCell
    Dim viableNeighborsColocatedFlag() As Boolean
    Dim viableNeighborsCount As Long
    
    Dim insertIdx As Long, tmpIdx As Long


    ' ---------------------------------------------------------------------------------------
    ' Begin - Create initial list of viable neighbor cells - cells within a given maximum distance
    ' ---------------------------------------------------------------------------------------




    ' Begin - Pre-calculate distance vectors from solution to each cell site (colo or sector split)
    viableNeighborsCount = 0

    For i = LBound(cellList) To UBound(cellList)

        path = EquirectangularPath(sLat, sLong, cellList(i).Lat, cellList(i).Long)


        If path.distance <= maxNeighborDistance + EPSILON And path.distance > CELL_MIN_DISTANCE Then ' we look at min distance
            viableNeighborsCount = viableNeighborsCount + 1
            
            'Debug.Print "Viable Neighbor: " & cellList(i).Name & " (" & cellList(i).UID & ")" & " - " & Format(path.distance, "0.00") & "mi"


            ReDim Preserve viableNeighbors(1 To viableNeighborsCount) As PossibleNeighborCell
            
            
            
            ' Insert into viableNeighbors() so that viableNeighbors() is sorted by distance
            insertIdx = viableNeighborsCount ' default: last item
            
            For j = 1 To viableNeighborsCount - 1
                If path.distance < viableNeighbors(j).path.distance Then
                    insertIdx = j
                    Exit For
                End If
            Next j
            
            ' Shift neighbors with greater distance to the right side of the array
            For j = viableNeighborsCount To insertIdx + 1 Step -1
                viableNeighbors(j) = viableNeighbors(j - 1)
            Next j
            
            
            
            viableNeighbors(insertIdx).neighbor_cell.cell = cellList(i)
            viableNeighbors(insertIdx).is_neighbor = False
            viableNeighbors(insertIdx).path = path
            
            'viableNeighbors(viableNeighborsCount).neighbor_cell.cell = cellList(i)
            'viableNeighbors(viableNeighborsCount).is_neighbor = False
            'viableNeighbors(viableNeighborsCount).path = path
            
            
        End If
        'End If
    Next i
    ' End - Pre-calculate distance vectors from solution to each cell site



    ' No 1st tier neighbors -> return empty list
    If viableNeighborsCount = 0 Then Exit Function

    

    ' Sort viable neighbors by distance using indices - DEPRECATED. Neighbors are pre-sorted
    'Dim neighbors_sortedIdx() As Long, tmpIdx As Long
    'ReDim neighbors_sortedIdx(1 To viableNeighborsCount) As Long

    'For i = 1 To viableNeighborsCount: neighbors_sortedIdx(i) = i: Next i

    'For i = 1 To viableNeighborsCount
    '    For j = i + 1 To viableNeighborsCount
    '        If viableNeighbors(neighbors_sortedIdx(j)).path.distance < viableNeighbors(neighbors_sortedIdx(i)).path.distance Then
    '            tmpIdx = neighbors_sortedIdx(i)
    '            neighbors_sortedIdx(i) = neighbors_sortedIdx(j)
    '            neighbors_sortedIdx(j) = tmpIdx
    '        End If
    '    Next j
    'Next i


    ' ---------------------------------------------------------------------------------------
    ' End - Create initial list of viable neighbor cells - cells within a given maximum distance
    ' ---------------------------------------------------------------------------------------


    Calc1stTierNeighbors_AnalyzeSectors viableNeighbors, viableNeighborsCount


    ' Narrow down possible neighbors by distance, bearing, and relationship between other viable neighbors
    Calc1stTierNeighbors_RemoveViableNeighborsBySimiliarBearing viableNeighbors, viableNeighborsCount

    'Calc1stTierNeighbors_CalcOffloadSectors viableNeighbors, viableNeighborsCount
    
    
    ' Narrow down set of neighbors by using an expanding search radius
    Calc1stTierNeighbors_RemoveViableNeighborsByRadiusExpansion viableNeighbors, viableNeighborsCount





    Dim neighborIncludeInFinalList() As Boolean
    Dim neighborFinalSortedIdx() As Long
    Dim neighborSortNormalizedAngles() As Double
    ReDim neighborIncludeInFinalList(1 To viableNeighborsCount) As Boolean
    ReDim neighborFinalSortedIdx(1 To viableNeighborsCount) As Long
    ReDim neighborSortNormalizedAngles(1 To viableNeighborsCount) As Double




    Dim finalNeighborCount As Long, finalNeighborIdx As Long
    Dim finalNeighborList() As NeighborCell

    finalNeighborCount = 0
    finalNeighborIdx = 1

    For i = 1 To viableNeighborsCount
        ' determine if we are including neighbor in final list
        If viableNeighbors(i).is_neighbor = True Then
            neighborIncludeInFinalList(i) = IIf(includeOnlyNeighborsWithOffloadSectors, viableNeighbors(i).neighbor_cell.is_offload_alpha Or viableNeighbors(i).neighbor_cell.is_offload_beta Or viableNeighbors(i).neighbor_cell.is_offload_gamma, True)

            If neighborIncludeInFinalList(i) Then finalNeighborCount = finalNeighborCount + 1


            ' pre-calcuate normalized angle for sorting
            neighborSortNormalizedAngles(i) = Angle_Normalize360(viableNeighbors(i).path.angle)
        End If
    Next i
    
    
    
    Calc1stTierNeighbors_CalculateCellWeights viableNeighbors, neighborIncludeInFinalList, viableNeighborsCount
    
    
    
    ' Begin - Determine neighbor weighting
    
    ' Weight is determined by the relative distance from solution to each neighbor (closer is better)
    ' Weight is proportional to 1/distance [Normalized Range: 0% - 100%]

    'Dim neighborWeightNormalizationFactorInverse As Double
    
    
    'neighborWeightNormalizationFactorInverse = 0
    
    ' First, calculate the normalization factor inverse (equal to sum of the each distance inverse)
    'For nIdx = 1 To viableNeighborsCount
    '    If viableNeighbors(nIdx).is_neighbor = True And neighborIncludeInFinalList(nIdx) = True Then
    '        neighborWeightNormalizationFactorInverse = neighborWeightNormalizationFactorInverse + (1 / viableNeighbors(nIdx).path.distance)
    '    End If
    'Next nIdx

    ' Calculate normalized weight as inverse of distance x normalization factor
    'For nIdx = 1 To viableNeighborsCount
    '    If viableNeighbors(nIdx).is_neighbor = True And neighborIncludeInFinalList(nIdx) = True Then
    '        viableNeighbors(nIdx).neighbor_cell.cell_weight = 1 / (viableNeighbors(nIdx).path.distance * neighborWeightNormalizationFactorInverse)
    '    End If
    'Next nIdx
    ' End - Determine neighbor weighting


    ' Begin - Sort indices of viable neighbors by their angle (optional - makes it easier to verify in google earth)
    For i = 1 To viableNeighborsCount: neighborFinalSortedIdx(i) = i: Next i

        For i = 1 To viableNeighborsCount
            For j = i + 1 To viableNeighborsCount
                If neighborIncludeInFinalList(neighborFinalSortedIdx(j)) Then
                    If neighborSortNormalizedAngles(neighborFinalSortedIdx(j)) < neighborSortNormalizedAngles(neighborFinalSortedIdx(i)) Then
                        tmpIdx = neighborFinalSortedIdx(j)
                        neighborFinalSortedIdx(j) = neighborFinalSortedIdx(i)
                        neighborFinalSortedIdx(i) = tmpIdx
                    End If
                End If
            Next j
    Next i
    ' End - Sort indices of viable neighbors by their angle (optional - makes it easier to verify in google earth)
    
    
    '---------------------------
    '
    Dim neighborDirectionalityVectX As Double, neighborDirectionalityVectY As Double
    
    neighborDirectionalityVectX = 0
    neighborDirectionalityVectY = 0
    
    For nIdx = 1 To viableNeighborsCount
        If viableNeighbors(nIdx).is_neighbor = True And neighborIncludeInFinalList(nIdx) = True Then
            Dim vecMag As Double
            
            vecMag = 1 / (viableNeighbors(nIdx).path.distance * viableNeighbors(nIdx).path.distance)
            'vecMag = viableNeighbors(nIdx).neighbor_cell.cell_weight
            
            neighborDirectionalityVectX = neighborDirectionalityVectX + vecMag * Cos(viableNeighbors(nIdx).path.angle)
            neighborDirectionalityVectY = neighborDirectionalityVectY + vecMag * Sin(viableNeighbors(nIdx).path.angle)
        End If
    Next nIdx
    
    Dim tmpMag As Double, tmpAngle As Double
    
    tmpMag = Sqr(neighborDirectionalityVectX * neighborDirectionalityVectX + neighborDirectionalityVectY * neighborDirectionalityVectY)
    tmpAngle = Atan2(neighborDirectionalityVectX, neighborDirectionalityVectY) * 180 / PI
    
    Debug.Print " Directionality: " & _
        Format(tmpMag, "0.00") & _
        " @ " & Format(tmpAngle, "0") & "°:" & Format(ModDouble(450 - tmpAngle, 360), "0") & "°H:"
    
    '---------------------------
    
    


    ReDim finalNeighborList(1 To finalNeighborCount) As NeighborCell

    For i = 1 To viableNeighborsCount
        nIdx = neighborFinalSortedIdx(i)

        If viableNeighbors(nIdx).is_neighbor = True Then


            If neighborIncludeInFinalList(nIdx) Then
                finalNeighborList(finalNeighborIdx) = viableNeighbors(nIdx).neighbor_cell
                
                finalNeighborList(finalNeighborIdx).distance = viableNeighbors(nIdx).path.distance
            
                'finalNeighborList(finalNeighborIdx).cell = viableNeighbors(nIdx).neighbor_cell.cell
                'finalNeighborList(finalNeighborIdx).cell_weight = neighborWeight(nIdx)
                
                'finalNeighborList(finalNeighborIdx).is_offload_alpha = viableNeighbors(nIdx).neighbor_cell.is_offload_alpha
                'finalNeighborList(finalNeighborIdx).is_offload_beta = viableNeighbors(nIdx).neighbor_cell.is_offload_beta
                'finalNeighborList(finalNeighborIdx).is_offload_gamma = viableNeighbors(nIdx).neighbor_cell.is_offload_gamma
                
                'finalNeighborList(finalNeighborIdx).sector_weight_alpha = viableNeighbors(nIdx).neighbor_cell.sector_weight_alpha
                'finalNeighborList(finalNeighborIdx).sector_weight_beta = viableNeighbors(nIdx).neighbor_cell.sector_weight_beta
                'finalNeighborList(finalNeighborIdx).sector_weight_gamma = viableNeighbors(nIdx).neighbor_cell.sector_weight_gamma

                finalNeighborIdx = finalNeighborIdx + 1
            End If
        End If
    Next i





    Calc1stTierNeighbors = finalNeighborList

    

    ' --------------------------------------------------------------------------------------------------------------------
    ' Begin - Print offload sectors - can be removed
    ' --------------------------------------------------------------------------------------------------------------------
    If DBG_PRINT_FLAG Then
        Debug.Print "--------------------"

        Dim offloadSectors As String
        Dim offloadSectorNames As String
        Dim is_offload As Boolean

        Dim sectorName As String


        For i = 1 To finalNeighborCount
            offloadSectors = ""
            offloadSectorNames = ""

            For j = 1 To 3
                is_offload = Choose(j, finalNeighborList(i).is_offload_alpha, finalNeighborList(i).is_offload_beta, finalNeighborList(i).is_offload_gamma)

                If is_offload Then
                    sectorName = Choose(j, "Alpha", "Beta", "Gamma")

                    offloadSectors = offloadSectors & "," & j
                    offloadSectorNames = offloadSectorNames & "," & sectorName
                End If
            Next j

            offloadSectors = Mid(offloadSectors, 2)
            offloadSectorNames = Mid(offloadSectorNames, 2)

            Debug.Print finalNeighborList(i).cell.Name & " " & offloadSectorNames & " (" & finalNeighborList(i).cell.UID & "-" & offloadSectors & ")"
        Next i


        Dim offloadSectorsAll As String

        For i = 1 To finalNeighborCount
            offloadSectors = ""
            offloadSectorNames = ""

            For j = 1 To 3
                is_offload = Choose(j, finalNeighborList(i).is_offload_alpha, finalNeighborList(i).is_offload_beta, finalNeighborList(i).is_offload_gamma)

                If is_offload Then offloadSectors = offloadSectors & "," & j
            Next j

            offloadSectors = Mid(offloadSectors, 2)
            offloadSectorNames = Mid(offloadSectorNames, 2)

            offloadSectorsAll = offloadSectorsAll & finalNeighborList(i).cell.UID & "-" & offloadSectors & "; "
        Next i

        Debug.Print "--------------------"
        Debug.Print offloadSectorsAll
    End If
    ' --------------------------------------------------------------------------------------------------------------------
    ' End - Print offload sectors - can be removed
    ' --------------------------------------------------------------------------------------------------------------------



End Function

Private Function Cells_AreColocated(ByRef cell1 As CellSite, ByRef cell2 As CellSite) As Boolean

    Cells_AreColocated = Abs(cell1.Lat - cell2.Lat) <= EPSILON And Abs(cell1.Long - cell2.Long) <= EPSILON

End Function
Private Sub Calc1stTierNeighbors_RemoveViableNeighborsBySimiliarBearing(ByRef viableNeighbors() As PossibleNeighborCell, viableNeighborsCount As Long)

    Debug.Assert LBound(viableNeighbors) = 1
    Debug.Assert UBound(viableNeighbors) = viableNeighborsCount
    'Debug.Assert LBound(viableNeighbors) = LBound(neighbors_sortedIdx)
    'Debug.Assert UBound(viableNeighbors) = UBound(neighbors_sortedIdx)
    
    
    Const CELL_EFFECTIVE_RADIUS = 1 ' in miles, - a highly conservative estimate of a cell's effective radius
    
    ' in miles, the minimum radial distance from solution location between two neighbor cells for one of the cells to disqualify the other as a neighbor
    Const SAME_BEARING_MIN_SEPARATION_ANGLE_DIST = 0.2
    

    Dim i As Long, j As Long
    Dim nIdx As Long, pIdx As Long

    ' --------------------------------------------------------------------------------------
    ' Begin - Narrow down possible neighbors by distance, bearing, and relationship between other viable neighbors
    ' ---------------------------------------------------------------------------------------
    Dim anyNeighborConflicts As Boolean
    Dim neighborHasSimiliarBearing As Boolean, neighborIsAtSameLocation As Boolean, neighborIsInCloserCellsEffectiveCellRadius As Boolean
    Dim isInEffectiveCellRadius As Boolean



    For i = 1 To viableNeighborsCount
        nIdx = i 'neighbors_sortedIdx(i)

        If DBG_PRINT_FLAG Or True Then Debug.Print "Evaluating: " & nIdx & ":" & viableNeighbors(nIdx).neighbor_cell.cell.Name & " (" & viableNeighbors(nIdx).neighbor_cell.cell.UID & ")" & " - " & Format(viableNeighbors(nIdx).path.angle, "0") & "°/" & Format(viableNeighbors(nIdx).path.distance, "0.00") & "mi"


        ' --------------------------------------------------------------------------------------
        ' Begin - Disqualify viable neighbor if a closer neighbor exists which has a similiar bearing
        ' ---------------------------------------------------------------------------------------
        anyNeighborConflicts = False

        ' check if within  neighbor cell's effective radius. If yes, we automatically assume it's a neighbor
        isInEffectiveCellRadius = viableNeighbors(nIdx).path.distance <= CELL_EFFECTIVE_RADIUS

        
        'If isInEffectiveCellRadius = False Then
            For j = 1 To i - 1 ' Loop through previous neighbors (in sorted order)
                pIdx = j 'neighbors_sortedIdx(j)
                
                ' check if different cell but same location (e.g., colo or sector split). if yes, we will not evaluate its bearing
                neighborIsAtSameLocation = Cells_AreColocated(viableNeighbors(nIdx).neighbor_cell.cell, viableNeighbors(pIdx).neighbor_cell.cell)
                
                ' check if within previous closer cell's effective radius. If yes, we do not consider this cell to be a neighbor
                neighborIsInCloserCellsEffectiveCellRadius = viableNeighbors(pIdx).path.distance <= CELL_EFFECTIVE_RADIUS
                
    
                If neighborIsAtSameLocation Then
                    If DBG_PRINT_FLAG Or True Then Debug.Print "- SKIP: " & pIdx & ":" & viableNeighbors(pIdx).neighbor_cell.cell.Name & " (Same cell)"
                ElseIf neighborIsAtSameLocation = False And neighborIsInCloserCellsEffectiveCellRadius Then
                    If DBG_PRINT_FLAG Or True Then Debug.Print "- NOK: " & pIdx & ":" & viableNeighbors(pIdx).neighbor_cell.cell.Name & " (In previous cell's radius)"
                'ElseIf viableNeighbors(pIdx).is_neighbor = False Then
                '    If DBG_PRINT_FLAG Or True Then Debug.Print "- SKIP: " & pIdx & ":" & viableNeighbors(pIdx).neighbor_cell.cell.Name & " (Not a neighbor)"
                End If
    
    
    
                If neighborIsAtSameLocation = False Then  'And viableNeighbors(pIdx).is_neighbor Then
                
                
                
                
    
                
                    'sepAngleDist = viableNeighbors(pIdx).path.distance
                    'If sepAngleDist < CELL_EFFECTIVE_RADIUS Then CELL_EFFECTIVE_RADIUS
    
                    'minSepAngle = Atn(CELL_EFFECTIVE_RADIUS / viableNeighbors(pIdx).path.distance) * 180 / PI
    
    
                    neighborHasSimiliarBearing = False
    
                    ' Check if closer cell is very close to the current neighbor we are evaluating (distance-wise). If so, then we do not
                    ' look to see if the previous cell has a similiar bearing
                    If Abs(viableNeighbors(pIdx).path.distance - viableNeighbors(nIdx).path.distance) > SAME_BEARING_MIN_SEPARATION_ANGLE_DIST Then
                        
                        
                        
                        '---------------------------------------
                        Dim minSepAngle As Double, angleDiff As Double
                        
                        minSepAngle = Atn(CELL_EFFECTIVE_RADIUS / viableNeighbors(pIdx).path.distance) * 180 / PI
                        
                        angleDiff = Angle_SmallestDifference(viableNeighbors(nIdx).path.angle, viableNeighbors(pIdx).path.angle)
                        
                        If angleDiff < minSepAngle Then neighborHasSimiliarBearing = True
                        
                        If DBG_PRINT_FLAG Or True Then Debug.Print "- " & IIf(neighborHasSimiliarBearing, "NOK", "OK ") & " : " & pIdx & ":" & viableNeighbors(pIdx).neighbor_cell.cell.Name & " (" & viableNeighbors(pIdx).neighbor_cell.cell.UID & ")" & " - " & Format(viableNeighbors(pIdx).path.angle, "0") & "°/" & Format(viableNeighbors(pIdx).path.distance, "0.00") & "mi (Min: " & Format(minSepAngle, "0") & "°, Diff: " & Format(angleDiff, "0") & "°)"

                        '---------------------------------------
                        'Dim lineK As Double, lineKSquaredPlusOne As Double
                        'Dim minSepAngle As Double, tanPtX As Double, tanPtY As Double
                        
                        
                        'If neighborIsInCloserCellsEffectiveCellRadius = True Then
                        '    anyNeighborConflicts = True
                        '    Exit For
                        'End If
                        
                        ' Find the min separation angle, which is the angle between a line segment from source location
                        '   which is also tangent to neighbor's effective radius circle.
                        '       y = mx, y^2 + (x - d)^2 = R^2 (d = neighbor distance, R = effective radius)
                        '   We want only one solution (tangent), so we need a double root for m in the quadratic (i.e. det = b^2 - 4ac = 0)
                        '   Solving for m: m = Sqrt((d/R)^2 - 1) which only exists if d > R (neighborIsInCloserCellsEffectiveCellRadius=False)
                        '    (the case where d <= R is handled already)
                        '
                        'Debug.Assert viableNeighbors(pIdx).path.distance > CELL_EFFECTIVE_RADIUS ' Just in case
                        
                        'lineKSquaredPlusOne = (viableNeighbors(pIdx).path.distance / CELL_EFFECTIVE_RADIUS) ^ 2
                        'lineSlope = Sqr(lineKSquaredPlusOne - 1) ' We take the positive root
                        
                        ' Now find the point of tangency using formulas above
                        'tanPtX = viableNeighbors(pIdx).path.distance / lineKSquaredPlusOne
                        'tanPtY = lineSlope * tanPtX
                        
                        'minSepAngle = Atn(tanPtY / tanPtX) * 180 / PI
                        

                        'angleDiff = Angle_SmallestDifference(viableNeighbors(nIdx).path.angle, viableNeighbors(pIdx).path.angle)
    
                        'If angleDiff < minSepAngle Then neighborHasSimiliarBearing = True
                        
                        'If DBG_PRINT_FLAG Or True Then Debug.Print "- " & IIf(neighborHasSimiliarBearing, "NOK", "OK ") & " : " & pIdx & ":" & viableNeighbors(pIdx).neighbor_cell.cell.Name & " (" & viableNeighbors(pIdx).neighbor_cell.cell.UID & ")" & " - " & Format(viableNeighbors(pIdx).path.angle, "0") & "°/" & Format(viableNeighbors(pIdx).path.distance, "0.00") & "mi (Min: " & Format(minSepAngle, "0") & "°, Diff: " & Format(angleDiff, "0") & "°)"
                        ' -------------------------------------

                        ' --------------------------------------
                        ' Find if line segment between source location and current neighbor intersects a circle created
                        ' within the effective radius of a previous (closer) neighbor. If so, then we consider the previous neighbor
                        ' to have a similiar bearing and therefore we do not consider this cell as a neighbor
                        'Dim y1 As Double, x1 As Double
                        'Dim dx As Double, dy As Double
                        'Dim cx As Double, cy As Double
                        'Dim A As Double, B As Double, C As Double
                        'Dim det As Double, t As Double
                        
                        'x1 = 0
                        'y1 = 0
                        't = 0
                        'dx = viableNeighbors(nIdx).path.dx
                        'dy = viableNeighbors(nIdx).path.dy
                        'cx = viableNeighbors(pIdx).path.dx
                        'cy = viableNeighbors(pIdx).path.dy
                    
                        'A = dx * dx + dy * dy
                        'B = 2 * (dx * (x1 - cx) + dy * (y1 - cy))
                        'C = (x1 - cx) * (x1 - cx) + (y1 - cy) * (y1 - cy) - CELL_EFFECTIVE_RADIUS * CELL_EFFECTIVE_RADIUS
                        
                        'det = B * B - 4 * A * C
                        
                        'If det > 0 Then
                        '    ' Two solutions (line intersects circle)
                        '
                        '    ' Need need to check if the line *segment* intersects the circle (0 <= t <= 1)
                        '    t = (-B + Sqr(det)) / (2 * A)
                        '    If t >= 0 And t <= 1 Then neighborHasSimiliarBearing = True
                        '
                        '    t = (-B - Sqr(det)) / (2 * A)
                        '    If t >= 0 And t <= 1 Then neighborHasSimiliarBearing = True
                        'ElseIf det = 0 Then
                        '    ' One solution (line is tangent to circle)
                        '    t = -B / (2 * A)
                        '    If t >= 0 And t <= 1 Then neighborHasSimiliarBearing = True
                        'ElseIf det < 0 Then
                            ' No solution (no intersection)
                        'End If
                        
                        
                        'If DBG_PRINT_FLAG Or True Then Debug.Print "- " & IIf(neighborHasSimiliarBearing, "NOK", "OK ") & " : " & pIdx & ":" & viableNeighbors(pIdx).neighbor_cell.cell.Name & " (" & viableNeighbors(pIdx).neighbor_cell.cell.UID & ")" & " - " & Format(viableNeighbors(pIdx).path.angle, "0") & "°/" & Format(viableNeighbors(pIdx).path.distance, "0.00") & "mi (t: " & t & ")"

                        ' ------------------------------------------------------------------
    
    
    
                    Else
                        If DBG_PRINT_FLAG Or True Then Debug.Print "- SKIP: " & pIdx & ":" & viableNeighbors(pIdx).neighbor_cell.cell.Name & " (Distance equivalence)"
                    End If
    
    
    
                    If neighborHasSimiliarBearing Then
                        anyNeighborConflicts = True
                        Exit For
                    End If
    
                End If
    
            Next j
            
        'Else
        '    If DBG_PRINT_FLAG Or True Then Debug.Print "- Cell is within effective radius"
        'End If



        If anyNeighborConflicts = False Then viableNeighbors(nIdx).is_neighbor = True
        ' --------------------------------------------------------------------------------------
        ' End - Disqualify viable neighbor if a closer neighbor exists which has a similiar bearing
        ' ---------------------------------------------------------------------------------------
        


    Next i

    ' --------------------------------------------------------------------------------------
    ' End - Narrow down possible neighbors by distance, bearing, and relationship between other viable neighbors
    ' ---------------------------------------------------------------------------------------



End Sub
Private Sub Calc1stTierNeighbors_RemoveViableNeighborsByRadiusExpansion(ByRef viableNeighbors() As PossibleNeighborCell, viableNeighborsCount As Long, Optional expandIncrement As Double = 0.2)

    Debug.Assert LBound(viableNeighbors) = 1
    Debug.Assert UBound(viableNeighbors) = viableNeighborsCount
    'Debug.Assert LBound(viableNeighbors) = LBound(neighbors_sortedIdx)
    'Debug.Assert UBound(viableNeighbors) = UBound(neighbors_sortedIdx)
    
    Dim i As Long, j As Long
    Dim nIdx As Long
    


    
    Dim prevRadius As Double, searchRadius As Double, newRadius As Double
    Dim lastIdxInSearchRadius As Long
    
    Dim passNo As Long, numContiguousPassesWithoutNeighbors As Integer
    Dim numNeighborsThisPass As Long
    Dim numNeighborsLastPass As Long
    Dim numNeighborsTotal As Long

    
    ' Find closest neighbor - This will be our starting radius
    For i = 1 To viableNeighborsCount
        nIdx = i 'neighbors_sortedIdx(i)
        
        If viableNeighbors(nIdx).is_neighbor Then
            searchRadius = viableNeighbors(nIdx).path.distance
            lastIdxInSearchRadius = i
            Exit For
        End If
    Next i
    
    
    Debug.Print "Filter by Radius Expansion"
    
    
    For i = 1 To viableNeighborsCount
        nIdx = i 'neighbors_sortedIdx(i)
        
        If viableNeighbors(nIdx).is_neighbor Then
             Debug.Print "---" & i & ":" & nIdx & ":" & viableNeighbors(nIdx).neighbor_cell.cell.Name & " (" & viableNeighbors(nIdx).neighbor_cell.cell.UID & ")" & " - " & Format(viableNeighbors(nIdx).path.angle, "0") & "°/" & Format(viableNeighbors(nIdx).path.distance, "0.00") & "mi"
        End If
    Next i
    
    lastIdxInSearchRadius = 0
    passNo = 0
    numNeighborsTotal = 0
    numNeighborsLastPass = 0
    numContiguousPassesWithoutNeighbors = 0
    
    prevRadius = 0
    'searchRadius = expandIncrement
    
    Do While lastIdxInSearchRadius < viableNeighborsCount
        passNo = passNo + 1
        
        
        Debug.Print "- Spiral Pass #" & passNo & " (<= " & Format(searchRadius, "0.00") & "mi)"
        
        numNeighborsThisPass = 0
        
        For i = lastIdxInSearchRadius + 1 To viableNeighborsCount
            nIdx = i 'neighbors_sortedIdx(i)
            
            If viableNeighbors(nIdx).is_neighbor And viableNeighbors(nIdx).path.distance <= searchRadius Then
                lastIdxInSearchRadius = i
                'spiralRadius = viableNeighbors(nIdx).path.distance + SPIRAL_BASE_UNIT
                numNeighborsThisPass = numNeighborsThisPass + 1
            End If
                             
            If viableNeighbors(nIdx).is_neighbor Then
                Debug.Print "--- " & IIf(viableNeighbors(nIdx).path.distance <= searchRadius, "Y", "N") & " " & i & ":" & nIdx & ":" & viableNeighbors(nIdx).neighbor_cell.cell.Name & " (" & viableNeighbors(nIdx).neighbor_cell.cell.UID & ")" & " - " & Format(viableNeighbors(nIdx).path.angle, "0") & "°/" & Format(viableNeighbors(nIdx).path.distance, "0.00") & "mi"
            End If
            
        Next i
        
        numNeighborsTotal = numNeighborsTotal + numNeighborsThisPass
        
        If numNeighborsTotal > 0 Then
            If numNeighborsThisPass > 0 Then
                numContiguousPassesWithoutNeighbors = 0
            Else
                numContiguousPassesWithoutNeighbors = numContiguousPassesWithoutNeighbors + 1
            End If
            
            
            If passNo > 1 And numContiguousPassesWithoutNeighbors = 2 Then Exit Do
        End If
        
        'If numNeighborsThisPass = 0 And numNeighborsLastPass = 0 Then
        '    numContiguousPassesWithoutNeighbors = numContiguousPassesWithoutNeighbors + 1
        '
        '    If passNo > 1 And numContiguousPassesWithoutNeighbors = 1 Then Exit Do
        'End If
        
        
        
        numNeighborsLastPass = numNeighborsThisPass
        
        
        ' Expand search area
        newRadius = searchRadius + expandIncrement 'prevRadius + searchRadius
        prevRadius = searchRadius
        searchRadius = newRadius
    Loop
    
    
 
    For i = lastIdxInSearchRadius + 1 To viableNeighborsCount
        nIdx = i 'neighbors_sortedIdx(i)
        
        viableNeighbors(nIdx).is_neighbor = False
    Next i
    
    
End Sub
Private Function Calc1stTierNeighbors_AnalyzeSectors(ByRef viableNeighbors() As PossibleNeighborCell, viableNeighborsCount As Long)
    
    ' Calculates the following sector information for each viable neighbor cell (co-located cells are analyzed together)
    '   - Offload sectors (sectors facing toward the subject location)
    '   - Offload sector weighting (how close a specific sector is facing towards the subject location - smaller angle results in greater weight)
    '   - Sector balancing
    
    
    'in degrees, maximum difference between the secondary offload sector to the primary offload sector for the secondary to be considered as an additional offload sector
    Const MULTI_OFFLOAD_SECTOR_ANGLE_THRESHOLD = 30
    
    
    Const MAX_SECTORS_PER_CELL As Integer = 9 ' arbitrary maximum
    
    
    Dim i As Long, j As Long, k As Long
    Dim nIdx As Long, pIdx As Long
    
    
    Dim angleDiff1 As Double, angleDiff2 As Double
    Dim azimuth1 As Integer, azimuth2 As Integer
    'Dim sectorName1 As String, sectorName2 As String
    Dim tmpSector As Integer
    
    
    Dim sectorSortedIdx(1 To 3) As Integer
    Dim sectorAngleDiffs(1 To 3) As Double
    Dim sectorAzimuths(1 To 3) As Integer
    Dim sectorSorted(1 To 3) As Integer
    
    
    Dim viableNeighbors_alreadyAnalyzedSectorsFlag() As Boolean
    ReDim viableNeighbors_alreadyAnalyzedSectorsFlag(1 To viableNeighborsCount) As Boolean
    
    
    
    Dim possibleOffloadSectors_count As Long
    Dim possibleOffloadSectors_neighborIdx(1 To MAX_SECTORS_PER_CELL) As Long
    Dim possibleOffloadSectors_sector(1 To MAX_SECTORS_PER_CELL) As Integer
    Dim possibleOffloadSectors_angleDiff(1 To MAX_SECTORS_PER_CELL) As Double
    Dim possibleOffloadSectors_azimuth(1 To MAX_SECTORS_PER_CELL) As Integer
    Dim possibleOffloadSectors_sortedIdx(1 To MAX_SECTORS_PER_CELL) As Integer
    Dim possibleOffloadSectors_is_offload_sector(1 To MAX_SECTORS_PER_CELL) As Boolean
    
    
    Dim offloadSectorCount As Long
    
    'Dim possibleOffloadSectors_weights(1 To MAX_SECTORS_PER_CELL) As Double
    'Dim possibleOffloadSectors_weightNormalizationFactorInverse As Double
    
    Dim sectorAzimuth As Integer, sectorAngle As Double
    Dim sectorIdx As Integer
    Dim sectorName As String
    
    ' --------------------------------------------------------------------------------------------------------------------
    ' Begin - for each viable neighbor, create list of sectors in same location, and select the offload sectors
    ' --------------------------------------------------------------------------------------------------------------------
    For i = 1 To viableNeighborsCount
        nIdx = i 'sortedIdx(i)
    
        'viableNeighbors(nIdx).is_neighbor = True And
        If viableNeighbors_alreadyAnalyzedSectorsFlag(nIdx) = False Then
            If DBG_PRINT_FLAG Then Debug.Print viableNeighbors(nIdx).neighbor_cell.cell.Name & " (" & viableNeighbors(nIdx).neighbor_cell.cell.UID & ")" & " - " & Format(viableNeighbors(nIdx).path.angle, "0") & "° (" & Format(viableNeighbors(nIdx).path.distance, "0.00") & "mi)"
    
            possibleOffloadSectors_count = 0
    
            ' Begin - Add this neighbors sectors to possible offload sector list
            For j = 1 To 3
                sectorAzimuth = Choose(j, viableNeighbors(nIdx).neighbor_cell.cell.Alpha_Azimuth, viableNeighbors(nIdx).neighbor_cell.cell.Beta_Azimuth, viableNeighbors(nIdx).neighbor_cell.cell.Gamma_Azimuth)
    
                If sectorAzimuth <> AZIMUTH_UNDEFINED Then
                    possibleOffloadSectors_count = possibleOffloadSectors_count + 1
    
                    possibleOffloadSectors_neighborIdx(possibleOffloadSectors_count) = nIdx
                    possibleOffloadSectors_sector(possibleOffloadSectors_count) = j
                    possibleOffloadSectors_azimuth(possibleOffloadSectors_count) = sectorAzimuth
                    possibleOffloadSectors_is_offload_sector(possibleOffloadSectors_count) = False
    
                    ' Convert azimuths to angle (0-360 from North clockwise to +/- 0-180 counterclockwise from East)
                    sectorAngle = 90 - sectorAzimuth
                    ' Calculate angle difference between the angle of neighbor path and the angle of sector azimuth
                    possibleOffloadSectors_angleDiff(possibleOffloadSectors_count) = Angle_SmallestDifference(viableNeighbors(nIdx).path.angle + 180, sectorAngle)
                End If
            Next j
            ' End - Add this neighbors sectors to possible offload sector list
    
    
            ' Begin - Add colocated neighbor sectors to possible offload sector list
            For j = i + 1 To viableNeighborsCount
                pIdx = j 'sortedIdx(j)
    
                'viableNeighbors(pIdx).is_neighbor = True And
                If viableNeighbors_alreadyAnalyzedSectorsFlag(pIdx) = False Then
    
                    If Cells_AreColocated(viableNeighbors(nIdx).neighbor_cell.cell, viableNeighbors(pIdx).neighbor_cell.cell) Then
                        For k = 1 To 3
                            sectorAzimuth = Choose(k, viableNeighbors(pIdx).neighbor_cell.cell.Alpha_Azimuth, viableNeighbors(pIdx).neighbor_cell.cell.Beta_Azimuth, viableNeighbors(pIdx).neighbor_cell.cell.Gamma_Azimuth)
    
                            If sectorAzimuth <> AZIMUTH_UNDEFINED Then
                                possibleOffloadSectors_count = possibleOffloadSectors_count + 1
    
                                possibleOffloadSectors_neighborIdx(possibleOffloadSectors_count) = pIdx
                                possibleOffloadSectors_sector(possibleOffloadSectors_count) = k
                                possibleOffloadSectors_azimuth(possibleOffloadSectors_count) = sectorAzimuth
                                possibleOffloadSectors_is_offload_sector(possibleOffloadSectors_count) = False
    
                                ' Convert azimuths to angle (0-360 from North clockwise to +/- 0-180 counterclockwise from East)
                                sectorAngle = 90 - sectorAzimuth
                                ' Calculate angle difference between the angle of neighbor path and the angle of sector azimuth
                                possibleOffloadSectors_angleDiff(possibleOffloadSectors_count) = Angle_SmallestDifference(viableNeighbors(pIdx).path.angle + 180, sectorAngle)
                            End If
                        Next k
    
                        'viableNeighbors(pIdx).aleady_analyzed_sectors = True
                        viableNeighbors_alreadyAnalyzedSectorsFlag(pIdx) = True
                    End If
    
                End If
            Next j
            ' End - Add colocated neighbor sectors to possible offload sector list
    
            ' Begin - Sort possible offload sectors by the angle difference between each azmith and the neighor path.
            For j = 1 To possibleOffloadSectors_count: possibleOffloadSectors_sortedIdx(j) = j: Next j 'Prefill sorted array with indices 1,2,3,...,possibleOffloadSectors_count
    
    
            ' sort array  indirectly using the sector indices
            For j = 1 To possibleOffloadSectors_count
                For k = j + 1 To possibleOffloadSectors_count
                    If possibleOffloadSectors_angleDiff(possibleOffloadSectors_sortedIdx(k)) < possibleOffloadSectors_angleDiff(possibleOffloadSectors_sortedIdx(j)) Then
                        tmpSector = possibleOffloadSectors_sortedIdx(j)
                        possibleOffloadSectors_sortedIdx(j) = possibleOffloadSectors_sortedIdx(k)
                        possibleOffloadSectors_sortedIdx(k) = tmpSector
                    End If
                Next k
            Next j
            ' End - Sort possible offload sectors by the angle difference between each azmith and the neighor path.
    
            ' Now possibleOffloadSectors_sortedIdx() contains the sector indices in order of how close each sector
            ' is pointing in the direction of the solution
            '   possibleOffloadSectors_sortedIdx(1) is the sector index of the primary offload sector
            '   possibleOffloadSectors_sortedIdx(2) is the sector index of the seconday offload sector
    
    
    
    
    
            sectorIdx = possibleOffloadSectors_sortedIdx(1) ' Primary offload sector
            angleDiff1 = possibleOffloadSectors_angleDiff(sectorIdx)
            
            possibleOffloadSectors_is_offload_sector(sectorIdx) = True
            offloadSectorCount = 1
    
            Select Case possibleOffloadSectors_sector(sectorIdx)
                Case 1: viableNeighbors(possibleOffloadSectors_neighborIdx(sectorIdx)).neighbor_cell.is_offload_alpha = True
                Case 2: viableNeighbors(possibleOffloadSectors_neighborIdx(sectorIdx)).neighbor_cell.is_offload_beta = True
                Case 3: viableNeighbors(possibleOffloadSectors_neighborIdx(sectorIdx)).neighbor_cell.is_offload_gamma = True
            End Select
    
    
            If possibleOffloadSectors_count > 1 Then
                For j = 2 To possibleOffloadSectors_count
    
                    sectorIdx = possibleOffloadSectors_sortedIdx(j) ' Secondary, tertiary,... offload sector
                    angleDiff2 = possibleOffloadSectors_angleDiff(sectorIdx)
    
                    ' check to see if angle difference between first and second offload candidates is small. If so, the second offload candidate is an actual offload sector
                    If angleDiff2 - angleDiff1 <= MULTI_OFFLOAD_SECTOR_ANGLE_THRESHOLD Then
                        possibleOffloadSectors_is_offload_sector(sectorIdx) = True
                        offloadSectorCount = offloadSectorCount + 1
                    
                        Select Case possibleOffloadSectors_sector(sectorIdx)
                            Case 1: viableNeighbors(possibleOffloadSectors_neighborIdx(sectorIdx)).neighbor_cell.is_offload_alpha = True
                            Case 2: viableNeighbors(possibleOffloadSectors_neighborIdx(sectorIdx)).neighbor_cell.is_offload_beta = True
                            Case 3: viableNeighbors(possibleOffloadSectors_neighborIdx(sectorIdx)).neighbor_cell.is_offload_gamma = True
                        End Select
                    End If
    
                Next j
            End If
    
            ' ------------------------------------------------------------------
            ' Begin - Calculator sector weights
            ' ------------------------------------------------------------------
            
            ' Weight is determined by the relative angle difference of each offload sector to the solution (closer is better)
            ' Weight is proportional to <angle difference of all offload sectors LESS angle difference of sector>
            ' Weight is normalized [Normalized Range: 0% - 100%]
            
            Dim sectorWeightTotalAngleDiffs As Double
            Dim sectorWeightNormalizationFactor As Double
            
            
            'Dim sectorWeightNormalizationFactorInverse As Double
            Dim sectorWeight As Double
            'sectorWeightNormalizationFactorInverse = 0
            
            sectorWeightTotalAngleDiffs = 0
            sectorWeightNormalizationFactor = 0
            
            For sectorIdx = 1 To possibleOffloadSectors_count
                If possibleOffloadSectors_is_offload_sector(sectorIdx) Then
                    'sectorWeightNormalizationFactorInverse = sectorWeightNormalizationFactorInverse + (1 / possibleOffloadSectors_angleDiff(sectorIdx))
                    sectorWeightTotalAngleDiffs = sectorWeightTotalAngleDiffs + possibleOffloadSectors_angleDiff(sectorIdx)
                End If
            Next sectorIdx
            
            
            If offloadSectorCount > 1 Then sectorWeightNormalizationFactor = 1 / (sectorWeightTotalAngleDiffs * (offloadSectorCount - 1))
        
            
            For sectorIdx = 1 To possibleOffloadSectors_count
                If possibleOffloadSectors_is_offload_sector(sectorIdx) Then
                    sectorWeight = 1 'Special case
                    If offloadSectorCount > 1 Then sectorWeight = possibleOffloadSectors_angleDiff(sectorIdx) * sectorWeightNormalizationFactor
                    

                    Select Case possibleOffloadSectors_sector(sectorIdx)
                        Case 1: viableNeighbors(possibleOffloadSectors_neighborIdx(sectorIdx)).neighbor_cell.sector_weight_alpha = sectorWeight
                        Case 2: viableNeighbors(possibleOffloadSectors_neighborIdx(sectorIdx)).neighbor_cell.sector_weight_beta = sectorWeight
                        Case 3: viableNeighbors(possibleOffloadSectors_neighborIdx(sectorIdx)).neighbor_cell.sector_weight_gamma = sectorWeight
                    End Select
                End If
            Next sectorIdx
            ' ------------------------------------------------------------------
            ' End - Calculator sector weights
            ' ------------------------------------------------------------------
    
    
            ' ------------------------------------------------------------------
            ' Begin - Calculate coverage directivity vector (Work in progress - playing with a new idea)
            ' ------------------------------------------------------------------
            Dim covDirectivityVectorX As Double, covDirectivityVectorY As Double
            Dim covDirectivityMag As Double, covDirectivityAngle As Double
            
            covDirectivityVectorX = 0
            covDirectivityVectorY = 0
            
            For sectorIdx = 1 To possibleOffloadSectors_count
                covDirectivityVectorX = covDirectivityVectorX + Cos((90 - possibleOffloadSectors_azimuth(sectorIdx)) * PI / 180)
                covDirectivityVectorY = covDirectivityVectorY + Sin((90 - possibleOffloadSectors_azimuth(sectorIdx)) * PI / 180)
            Next sectorIdx
            
            covDirectivityMag = Sqr(covDirectivityVectorX * covDirectivityVectorX + covDirectivityVectorY * covDirectivityVectorY)
            covDirectivityAngle = Atan2(covDirectivityVectorX, covDirectivityVectorY) * 180 / PI
            
            viableNeighbors(nIdx).neighbor_cell.coverage_directivity_magnitude = covDirectivityMag
            viableNeighbors(nIdx).neighbor_cell.coverage_directivity_angle = covDirectivityAngle
            viableNeighbors(nIdx).neighbor_cell.coverage_directivity_magnitude_scaled = covDirectivityMag * possibleOffloadSectors_count
            viableNeighbors(nIdx).neighbor_cell.coverage_directivity_angle_diff = Angle_SmallestDifference(covDirectivityAngle, viableNeighbors(nIdx).path.angle)
            
            
            ' Now apply to co-located cells
            For j = i + 1 To viableNeighborsCount
                pIdx = j 'sortedIdx(j)
    
                If viableNeighbors_alreadyAnalyzedSectorsFlag(pIdx) = False Then
                    If Cells_AreColocated(viableNeighbors(nIdx).neighbor_cell.cell, viableNeighbors(pIdx).neighbor_cell.cell) Then
                        'Debug.Print viableNeighbors(pIdx).neighbor_cell.cell.Name, covDirectivityMag, covDirectivityAngle
                        
                        viableNeighbors(pIdx).neighbor_cell.coverage_directivity_magnitude = viableNeighbors(nIdx).neighbor_cell.coverage_directivity_magnitude
                        viableNeighbors(pIdx).neighbor_cell.coverage_directivity_angle = viableNeighbors(nIdx).neighbor_cell.coverage_directivity_angle
                        viableNeighbors(pIdx).neighbor_cell.coverage_directivity_magnitude_scaled = viableNeighbors(nIdx).neighbor_cell.coverage_directivity_magnitude_scaled
                        viableNeighbors(nIdx).neighbor_cell.coverage_directivity_angle_diff = viableNeighbors(pIdx).neighbor_cell.coverage_directivity_angle_diff
                    End If
                End If
            Next j
            ' ------------------------------------------------------------------
            ' End - Calculate coverage directivity vector
            ' ------------------------------------------------------------------
            
            
    
            ' Print offload sectors
            If DBG_PRINT_FLAG Then
                For j = 1 To possibleOffloadSectors_count
                    Dim is_offload As Boolean
    
                    sectorIdx = possibleOffloadSectors_sortedIdx(j)
                    sectorName = Choose(possibleOffloadSectors_sector(sectorIdx), "Alpha", "Beta", "Gamma")
                    is_offload = Choose(possibleOffloadSectors_sector(sectorIdx), _
                            viableNeighbors(possibleOffloadSectors_neighborIdx(sectorIdx)).neighbor_cell.is_offload_alpha, _
                            viableNeighbors(possibleOffloadSectors_neighborIdx(sectorIdx)).neighbor_cell.is_offload_beta, _
                            viableNeighbors(possibleOffloadSectors_neighborIdx(sectorIdx)).neighbor_cell.is_offload_gamma _
                        )
                        
                    Debug.Print "- " & viableNeighbors(possibleOffloadSectors_neighborIdx(sectorIdx)).neighbor_cell.cell.Name & " " & sectorName & " (" & possibleOffloadSectors_azimuth(sectorIdx) & ", " & Angle_Normalize180(90 - possibleOffloadSectors_azimuth(sectorIdx)) & "°, Diff: " & Format(possibleOffloadSectors_angleDiff(sectorIdx), "0") & ")" & IIf(is_offload, "*", "")
                Next j
            End If
    
        End If
    Next i
    ' --------------------------------------------------------------------------------------------------------------------
    ' End - for each viable neighbor, create list of sectors in same location, and select the target offload sectors
    ' --------------------------------------------------------------------------------------------------------------------



End Function
Private Sub Calc1stTierNeighbors_CalculateCellWeights(ByRef viableNeighbors() As PossibleNeighborCell, neighborIncludeInFinalList() As Boolean, viableNeighborsCount As Long)

    Debug.Assert LBound(viableNeighbors) = 1
    Debug.Assert UBound(viableNeighbors) = viableNeighborsCount
    Debug.Assert LBound(viableNeighbors) = LBound(neighborIncludeInFinalList)
    Debug.Assert UBound(viableNeighbors) = UBound(neighborIncludeInFinalList)
    
    Dim i As Long, j As Long
    Dim nIdx As Long
    

    ' --------------------------------------------------------------------
    ' Begin - Calculate neighbor cell weighting
    ' --------------------------------------------------------------------
    
    ' Each niehgbor cell weight is:
    '    - Determined by the relative distance from solution to each neighbor (closer is better)
    '    - Proportional to <total distance between solution and all neigbor cells LESS distance of solution to neighbor cell>
    '    - Normalized to 0% - 100%

    Dim finalNeighborCount As Long
    Dim finalNeighborTotalDistances As Double
    Dim finalNeighborWeightNormalizationFactor As Double
    
    finalNeighborTotalDistances = 0
    
    ' First, calculate the total distances of all final neighbors
    For nIdx = 1 To viableNeighborsCount
        If viableNeighbors(nIdx).is_neighbor = True And neighborIncludeInFinalList(nIdx) = True Then
            finalNeighborCount = finalNeighborCount + 1
            finalNeighborTotalDistances = finalNeighborTotalDistances + viableNeighbors(nIdx).path.distance
        End If
    Next nIdx
    
    ' Calculate normalization factor (Number of neighbors > 1)
    If finalNeighborCount > 1 Then finalNeighborWeightNormalizationFactor = 1 / (finalNeighborTotalDistances * (finalNeighborCount - 1))

    ' Second, calculate each neighbor's weight and then normalize it
    For nIdx = 1 To viableNeighborsCount
        If viableNeighbors(nIdx).is_neighbor = True And neighborIncludeInFinalList(nIdx) = True Then
            
            If finalNeighborCount = 1 Then ' Special case
                viableNeighbors(nIdx).neighbor_cell.cell_weight = 1
                Exit For
            End If
            
            viableNeighbors(nIdx).neighbor_cell.cell_weight = (finalNeighborTotalDistances - viableNeighbors(nIdx).path.distance)
            viableNeighbors(nIdx).neighbor_cell.cell_weight = viableNeighbors(nIdx).neighbor_cell.cell_weight * finalNeighborWeightNormalizationFactor
        End If
    Next nIdx
    ' --------------------------------------------------------------------
    ' End - Calculate neighbor cell weighting
    ' --------------------------------------------------------------------
    

End Sub

' Returns distance (in miles) and angle between two lat/long pairs.
'
' Uses a equirectangular plane approximation to the surface of a sphere near the given lat/lon. As a result, is it
' a fast computation but is valid for only lat/long pairs which are close to each other (<15mi apart)
Private Function EquirectangularPath(lat1 As Double, lon1 As Double, lat2 As Double, lon2 As Double) As Path_2D


    Dim lat1_rad As Double
    Dim lon1_rad As Double
    Dim lat2_rad As Double
    Dim lon2_rad As Double
    Dim dlat As Double
    Dim dlon As Double
    Dim dist As Double

    Dim dx As Double, dy As Double

    Dim ret As Path_2D

    ' Decimal degrees to radians
    lat1_rad = lat1 * PI / 180
    lon1_rad = lon1 * PI / 180
    lat2_rad = lat2 * PI / 180
    lon2_rad = lon2 * PI / 180

    dlat = lat2_rad - lat1_rad
    dlon = lon2_rad - lon1_rad

    ' calculate change in cartesian coordinates (units: radians)
    dx = dlon * Cos(lat1_rad) 'Cos((lat1_rad + lat2_rad) / 2)
    dy = dlat


    ret.dx = dx
    ret.dy = dy
    ret.angle = (Atan2(dx, dy) * (180 / PI))
    
    ret.distance = Sqr(dx * dx + dy * dy) ' distance in radians
    ret.distance = ret.distance * EARTH_RADIUS_MI ' convert radians -> miles


    EquirectangularPath = ret

End Function
' Converts any angle such so -180 <= angle < 180
Public Function Angle_Normalize180(ByVal angle As Double) As Double

    angle = ModDouble(angle, 360) ' reduce angle
    angle = ModDouble(angle + 360, 360)  ' force angle to be positive remainder so 0 <= angle <= 360

    If angle > 180 Then angle = angle - 360

    Angle_Normalize180 = angle

End Function
' Converts any angle such that 0 <= angle < 360
Private Function Angle_Normalize360(ByVal angle As Double) As Double

    angle = ModDouble(angle, 360) ' reduce angle
    angle = ModDouble(angle + 360, 360)  ' force angle to be positive remainder so 0 <= angle <= 360

    Angle_Normalize360 = angle

End Function
' Computes the difference between two angles
Private Function Angle_SmallestDifference(ByVal angle1 As Double, ByVal angle2 As Double, Optional signed As Boolean = False) As Double

    Dim angleDiff As Double
    
    If angle1 > 180 Or angle1 < -180 Then angle1 = Angle_Normalize180(angle1)
    If angle2 > 180 Or angle2 < -180 Then angle2 = Angle_Normalize180(angle2)

    angleDiff = angle2 - angle1
    angleDiff = ModDouble(angleDiff + 180, 360) - 180

    If signed = False Then angleDiff = Abs(angleDiff)

    Angle_SmallestDifference = angleDiff

End Function
' Replacement Modulus (division remainder) operator for double data types
Private Function ModDouble(ByVal numerator As Double, ByVal denominator As Double) As Double
    ModDouble = numerator - denominator * Int(numerator / denominator)
End Function
Private Function Atan2(ByVal x As Double, ByVal y As Double) As Double
    Dim theta As Double

    Const EPSILON = 0.0000001

    If (Abs(x) < EPSILON) Then
        If (Abs(y) < EPSILON) Then
            theta = 0#
        ElseIf (y > 0#) Then
            theta = PI / 2
        Else
            theta = -PI / 2
        End If
    Else
        theta = Atn(y / x)

        If (x < 0) Then
            If (y >= 0#) Then
                theta = PI + theta
            Else
                theta = theta - PI
            End If
        End If
    End If

    Atan2 = theta
End Function


Function NeighborCell_GenerateClusterDefStr(nCell As NeighborCell) As String

    Dim i As Long

    Dim sectors As String
    Dim is_offload As Boolean
    
    sectors = ""
    
    For i = 1 To 3
        is_offload = Choose(i, nCell.is_offload_alpha, nCell.is_offload_beta, nCell.is_offload_gamma)
        
        If is_offload Then sectors = sectors & "," & i
    Next i
    
    sectors = Mid$(sectors, 2)
        
    NeighborCell_GenerateClusterDefStr = nCell.cell.UID & "-" & sectors

End Function
' Converts neighborList array to cluster definition
Function NeighborCells_GenerateClusterDefStr(neighborList() As NeighborCell) As String

    On Error Resume Next

    Dim i As Long, j As Long

    Dim offloadSectorsAll As String
    
    
    For i = LBound(neighborList) To UBound(neighborList)
        offloadSectorsAll = offloadSectorsAll & NeighborCell_GenerateClusterDefStr(neighborList(i)) & "; "
    Next i
        
    NeighborCells_GenerateClusterDefStr = offloadSectorsAll

End Function

Private Function NeighborCells_CombineRelated_DRAFT()

    ' CA


    Exit Function

    
    Dim C  As Variant: C = Array(39.73875, -75.4541)

    ' North Percy: 39.962208, -75.153278
    ' Poplar St: 39.965778, -75.132778
    ' Monticello: 39.937361, -77.687464
    ' WIL PERSES: 39.73875, -75.4541
    
    
    Dim sLat As Double, sLong As Double
    Dim test2 As Variant

    
    Dim geoPlanDataFile As String, maxNeighborDistance As Double
    

    geoPlanDataFile = "N:\System Performance\Users\huntljo\Automation\PrePost Analysis\geoplan_report_1st_tier_neighbors.txt"
    maxNeighborDistance = 5
    
    
    If Dir(geoPlanDataFile) = "" Then
        MsgBox "Cell file not found:  " & vbCrLf & vbCrLf & geoPlanDataFile
        Exit Function
    End If
    
    

    sLat = CDbl(C(0))
    sLong = CDbl(C(1))


    Dim cellList_850() As CellSite
    Dim cellList_700() As CellSite

    cellList_850 = LoadCellListFromGeoPlanReport(geoPlanDataFile, "Cellular")
    cellList_700 = LoadCellListFromGeoPlanReport(geoPlanDataFile, "Upper 700 MHz")



    Dim neighbors_850() As NeighborCell
    Dim neighbors_700() As NeighborCell

    neighbors_850 = Calc1stTierNeighbors(cellList_850, sLat, sLong, maxNeighborDistance, True)
    neighbors_700 = Calc1stTierNeighbors(cellList_700, sLat, sLong, maxNeighborDistance, True)
    
    
    
    
    Dim neighborCellCollections() As Variant
    
    'neighborCellCollections = Array(neighbors_850, neighbors_700) - CANNOT COMPILE THIS LINE/.....

    Debug.Print


    Exit Function

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

End Function
