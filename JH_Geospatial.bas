Attribute VB_Name = "JH_Geospatial"
Option Explicit

' JH_Geospatial
' -----------------------------------------------------------------------------------------------------------------
' Common VBA functions used to calculate geospatial properties or similiar functions
'
' Joseph Huntley (joseph.huntley@vzw.com)
' ------------------------------------------------------------------------------------------------------------------
' 2016-01-26 (joe h)    - Migrated Google_XXX functions
' 2016-01-15 (joe h)    - Added ConvexHull, ConcaveHull algorithms
'                       - Added LineSegmentsIntersection, LineSegmentsIntersectionByPoint
'                       - Added OrientationByPoint
'                       - Added Points_CalcBoundingRect
' 2015-04-17 (joe h)    - Converted early bound XMLDOM objects to late bound versions
'

Private Const EARTH_RADIUS_MI = 3959
Private Const PI As Double = 3.14159265358979
Private Const EPSILON As Double = 0.000001 ' small number for comparing double values

Private Const FEET_PER_METER As Double = 3.28084


Private Type Vector_2D
    x As Double
    y As Double
    length As Double
    angle As Double
End Type

Type Point_2D
    ID As String
    
    x As Double
    y As Double
End Type

Type Rectangle_2D
    TopLeft As Point_2D
    BottomRight As Point_2D
End Type

Type Polygon_2D
    ID As String
    
    BoundingBox As Rectangle_2D
    
    Coordinates() As Point_2D
    CoordinateCount As Long
End Type

Public Type Geospatial_LineOfSight_Result
    ElevationStart As Double
    ElevationEnd As Double
    ElevationGain As Double
    
    HasLineOfSight As Boolean
    MinLineOfSightElevation As Double
    MaxPathElevation As Double
End Type


Public Type Google_GeoCode_Result
    status_code As String
    status_description As String

    street_number As String
    route_short As String
    route_long As String
    neighborhood As String
    city As String
    county As String
    state_short As String
    state_long As String
    postal_code As String
    postal_code_suffix As String
    
    geometry_lat As Double
    geometry_lon As Double
    location_type As String
End Type

Public Type Google_ElevationService_Result
    status_code As String
    status_description As String

End Type

Public Function Points_CalcBoundingRect(ByRef points() As Point_2D) As Rectangle_2D

    Dim rect As Rectangle_2D
    Dim i As Long
    
    rect.TopLeft.x = points(LBound(points)).x
    rect.TopLeft.y = points(LBound(points)).y
    
    rect.BottomRight.x = rect.TopLeft.x
    rect.BottomRight.y = rect.TopLeft.y
    
    For i = LBound(points) + 1 To UBound(points)
        If points(i).y < rect.TopLeft.y Then
            rect.TopLeft.y = points(i).y
        ElseIf points(i).y > rect.BottomRight.y Then
            rect.BottomRight.y = points(i).y
        End If
        
        If points(i).x < rect.TopLeft.x Then
            rect.TopLeft.x = points(i).x
        ElseIf points(i).x > rect.BottomRight.x Then
            rect.BottomRight.x = points(i).x
        End If
                
    Next i
    
    Points_CalcBoundingRect = rect

End Function

' To find orientation of ordered triplet (p, q, r).
' The function returns following values
' 0 --> p, q and r are colinear
' 1 --> Clockwise
' 2 --> Counterclockwise
Private Function OrientationByPoint(p As Point_2D, q As Point_2D, r As Point_2D) As Integer

    Dim val As Double
    
    val = (q.y - p.y) * (r.x - q.x) - (q.x - p.x) * (r.y - q.y)
 

    If Abs(val) < EPSILON Then
        OrientationByPoint = 0
    Else
        OrientationByPoint = IIf(val > 0, 1, 2) ' clock or counterclock wise
    End If
    
End Function


Public Function LineSegmentsIntersection(x1 As Double, y1 As Double, x2 As Double, y2 As Double, a1 As Double, b1 As Double, a2 As Double, b2 As Double, Optional inclusive As Boolean = True) As Boolean

    Dim dx As Double, dy As Double
    Dim da As Double, db As Double
    Dim s As Double, t As Double

    dx = x2 - x1
    dy = y2 - y1
    da = a2 - a1
    db = b2 - b1
    
    If (da * dy - db * dx) = 0 Then
        ' The segments are parallel.
        LineSegmentsIntersection = False
        Exit Function
    End If
    
    
    s = (dx * (b1 - y1) + dy * (x1 - a1)) / (da * dy - db * dx)
    t = (da * (y1 - b1) + db * (a1 - x1)) / (db * dx - da * dy)
    
    If inclusive Then
        LineSegmentsIntersection = (s >= 0 And s <= 1 And t >= 0 And t <= 1)
    Else
        LineSegmentsIntersection = (s > 0 And s < 1 And t > 0 And t < 1)
    End If

    ' If it exists, the point of intersection is:
    ' (x1 + t * dx, y1 + t * dy)
End Function
Public Function LineSegmentsIntersectionByPoint(p1 As Point_2D, q1 As Point_2D, p2 As Point_2D, q2 As Point_2D, Optional inclusive As Boolean = True) As Boolean

    LineSegmentsIntersectionByPoint = LineSegmentsIntersection(p1.x, p1.y, q1.x, q1.y, p2.x, p2.y, q2.x, q2.y, inclusive)

End Function


 
Private Function DistanceSquared(p1 As Point_2D, p2 As Point_2D) As Double
    DistanceSquared = (p1.x - p2.x) * (p1.x - p2.x) + (p1.y - p2.y) * (p1.y - p2.y)
End Function
Private Function PointOrientationSortCompare(p0 As Point_2D, p1 As Point_2D, p2 As Point_2D) As Integer

    Dim o As Integer
    Dim distSq_p0p2 As Double, distSq_p0p1 As Double
    
    o = OrientationByPoint(p0, p1, p2)
    
    If o <> 0 Then
        PointOrientationSortCompare = IIf(o = 2, -1, 1)
        Exit Function
    End If
    
    distSq_p0p1 = DistanceSquared(p0, p1)
    distSq_p0p2 = DistanceSquared(p0, p2)
    
    PointOrientationSortCompare = IIf(distSq_p0p2 >= distSq_p0p1, -1, 1)

End Function
' Concave Hull: Given a set of points, finds a polygon which encloses all the points (can be concave or convex)
'    Implementation of the k-th nearest neighbors algorithm described in http://repositorium.sdum.uminho.pt/bitstream/1822/6429/1/ConcaveHull_ACM_MYS.pdf
Public Function ConcaveHull(pointList() As Point_2D, Optional ByVal maxNeighbors As Integer = 3) As Polygon_2D

    'Dim data() As Variant: data = Worksheet_LoadColumnDataToArray(ThisWorkbook.Sheets("ConvexHullPoints").Cells, Array("SITE_NAME", "LATITUDE", "LONGITUDE"))
    'Dim pointList() As Point_2D, siteNamesByPoint() As String, dpL As Long, dpU As Long, dpC As Long, dpI As Long
    
    'dpL = LBound(data, 2): dpU = UBound(data, 2): dpC = dpU - dpL + 1
   
    'ReDim pointList(0 To dpC - 1) As Point_2D
    'ReDim siteNamesByPoint(0 To dpC - 1) As String
   
    'For dpI = 0 To dpC - 1
    '    siteNamesByPoint(dpI) = data(0, dpI + dpL)
    '    pointList(dpI).x = data(2, dpI + dpL) ' Longitude
    '    pointList(dpI).y = data(1, dpI + dpL) ' Latitude
    'Next dpI
        
    ' ---------------------------------------------------
    
    
    'Dim maxNeighbors As Integer: maxNeighbors = 3
    
    
    
    
    Dim i As Long, j As Long, k As Long
    Dim tmpIdx As Long, pIdx As Long
    Dim numPoints As Long
    
    Dim HullPolygon As Polygon_2D
    
    
    If maxNeighbors < 3 Then maxNeighbors = 3 ' Make sure maxNeighbors > 3
    
    Dim points() As Point_2D
    
    Dim points_oIdx() As Long       ' points_oIdx() will hold the the index of the original points within the input pointsList()
    Dim points_inHull() As Boolean  ' Tracks if point is in hull or not
    
    
    ' Begin - Initialize points_oIdx() with unique points
    numPoints = UBound(pointList) - LBound(pointList) + 1
    
    ReDim points_oIdx(0 To numPoints - 1) As Long
    
    
    Dim pointIsDuplicate As Boolean
    numPoints = 0
    
    For i = LBound(pointList) To UBound(pointList)
        pointIsDuplicate = False
        
        For j = LBound(pointList) To i - 1
            If Abs(pointList(i).x - pointList(j).x) < EPSILON And Abs(pointList(i).y - pointList(j).y) < EPSILON Then
                Debug.Print "Duplicate point: " & pointList(i).ID & " (" & pointList(j).ID & ")"
                pointIsDuplicate = True
                Exit For
            End If
        Next j
        
        If pointIsDuplicate = False Then
            numPoints = numPoints + 1
            points_oIdx(numPoints - 1) = i
        Else
        End If
    Next i
    
    ReDim Preserve points_oIdx(0 To numPoints - 1) As Long
    ' End - Initialize points_oIdx() with unique points
    
    
    ' Begin - Copy remaining unique points to new points() array
    ReDim points(0 To numPoints - 1) As Point_2D
        
    For i = 0 To numPoints - 1
        points(i) = pointList(points_oIdx(i))
    Next i
    ' Begin - Copy remaining unique points to new points() array
    
    
    If numPoints < 3 Then
        Err.Raise -1, , "Cannot compute concave hull for two or less unique points"
        Exit Function
    End If
    
    If numPoints = 3 Then
        ' The hull is equal to the three points remaining
        HullPolygon.Coordinates = points
        HullPolygon.CoordinateCount = 3
        
        HullPolygon.BoundingBox = Points_CalcBoundingRect(points)
        
        ConcaveHull = HullPolygon
        Exit Function
    End If
    
    
    ' Begin - Find index of point with lowest y-value from the original points() array
    Dim ymin_idx As Long
    
    ymin_idx = 0
    
    For i = 0 To numPoints - 1
        If ((points(i).y < points(ymin_idx).y) Or (points(i).y = points(ymin_idx).y And points(i).x < points(ymin_idx).x)) Then
            ymin_idx = i
        End If
    Next i
    ' End - Find index of point with lowest y-value from the original points() array

    
    
    
    Dim numPointsInHull As Long
    Dim hullPointsIdx() As Long
    
    ReDim hullPointsIdx(0 To numPoints - 1) As Long
    
    ReDim points_inHull(0 To numPoints - 1) As Boolean
    'ReDim points_hullOrder(0 To numPoints - 1) As Long
    
    
    Dim firstPointIdx As Long, curPointIdx As Long, prevPointIdx As Long, nextPointIdx As Long
    Dim prevAngle As Double
    
    Dim step As Long, attempts As Integer
    
    ' We have to run this algorithm until all points are enclosed by the hull. (We need to increase maxNeighbors each time)
    ' Ideally, we would only use one iteration. We use a GoTo Label to avoid recursion:
    attempts = 0
    
StartOver:
    attempts = attempts + 1
    
    If attempts = 200 Then
        Err.Raise -1, , "ConcaveHull: Number of attempts too high"
    End If

    Dim nearestNeighborsCount As Long
    Dim nearestNeighborsIdx() As Long, nearestNeighborsDistSq() As Double, nearestNeighborsAngle() As Double
    ReDim nearestNeighborsIdx(0 To maxNeighbors - 1) As Long
    ReDim nearestNeighborsDistSq(0 To maxNeighbors - 1) As Double
    ReDim nearestNeighborsAngle(0 To maxNeighbors - 1) As Double
    
    Dim nearestNeighborNextPointIdx As Long
    Dim isNearestNeighbor As Boolean
    Dim distSq As Double
    
    ' Initialize
    For i = 0 To numPoints - 1
        points_inHull(i) = False
        'points_hullOrder(i) = 0
    Next i
    
    
    ' Place bottom-most point at first position in hull
    numPointsInHull = 1
    firstPointIdx = ymin_idx
    points_inHull(firstPointIdx) = True
    'points_hullOrder(firstPointIdx) = numPointsInHull
    hullPointsIdx(0) = firstPointIdx
    
    curPointIdx = firstPointIdx
    prevPointIdx = -1
    prevAngle = 0
    
    step = 0
    
    ' Main loop - build hull one point at a time until we go back to the starting point
    Do While (curPointIdx <> firstPointIdx Or numPointsInHull = 1) 'And step < 20
        step = step + 1
        
        'Debug.Print siteNamesByPoint(points_oIdx(curPointIdx))
        
        ' Begin - Find the nearest neighbors to current point on the hull
        nearestNeighborsCount = 0
        
        For i = 0 To numPoints - 1
            ' Next posssible point cannot already be on the hull
            '   UNLESS it's the starting point and we've already built 3 hull points (prevents creating of a triangle excluding the rest of the pts)
            If points_inHull(i) = False Or (i = firstPointIdx And numPointsInHull > 3) Then
                distSq = DistanceSquared(points(curPointIdx), points(i))
                
                isNearestNeighbor = False
                
                ' Insert into max neighbors if array has NOT been filled OR it has been filled and the point we are evaluating is closer
                If nearestNeighborsCount < maxNeighbors Then
                    isNearestNeighbor = True
                ElseIf nearestNeighborsCount = maxNeighbors Then
                    isNearestNeighbor = distSq < nearestNeighborsDistSq(nearestNeighborsCount - 1)
                End If
                
                If isNearestNeighbor Then
                    Dim insertIdx As Long
                
                    ' Find index to insert
                    insertIdx = nearestNeighborsCount ' End of array plus 1 (default)
                    If insertIdx > maxNeighbors - 1 Then insertIdx = maxNeighbors - 1 ' Limit to maxNeighbors
                    
                    ' See if this point is closer than any other point. If so, we will displace the further points in the array (shift right)
                    For j = nearestNeighborsCount - 1 To 0 Step -1
                        If distSq < nearestNeighborsDistSq(j) Then insertIdx = j
                    Next j
                   
                    ' Displace further points in the array first
                    For j = nearestNeighborsCount To insertIdx + 1 Step -1
                        'k = j + 1
                        If j < maxNeighbors Then
                            nearestNeighborsIdx(j) = nearestNeighborsIdx(j - 1)
                            nearestNeighborsDistSq(j) = nearestNeighborsDistSq(j - 1)
                            nearestNeighborsAngle(j) = nearestNeighborsAngle(j - 1)
                        End If
                    Next j
                    
                    nearestNeighborsIdx(insertIdx) = i
                    nearestNeighborsDistSq(insertIdx) = distSq
                    
                    If nearestNeighborsCount < maxNeighbors Then nearestNeighborsCount = nearestNeighborsCount + 1
                    
                End If
                
                
            End If
        Next i
        ' End - Find the nearest neighbors to current point on the hull - nearestNeighborsIdx() is now sorted by distance
        
        
        Dim nearestNeighborsSortAngle() As Double
        ReDim nearestNeighborsSortAngle(0 To maxNeighbors - 1) As Double
        
        For i = 0 To nearestNeighborsCount - 1
            Dim angle As Double
            
            angle = (180 / PI) * Atan2( _
                points(nearestNeighborsIdx(i)).x - points(curPointIdx).x, _
                points(nearestNeighborsIdx(i)).y - points(curPointIdx).y)
            
            nearestNeighborsAngle(i) = Angle_Normalize180(90 - angle)
            nearestNeighborsSortAngle(i) = Angle_Normalize180(90 - (angle + prevAngle))
        Next i
        
        
        'Debug.Print "- Neighbors by Distance/Angle"
        'For i = 0 To nearestNeighborsCount - 1
        '    Debug.Print "   " & siteNamesByPoint(points_oIdx(nearestNeighborsIdx(i))) & " (" & Format(nearestNeighborsDistSq(i), "0.00e-#") & "," & Round(nearestNeighborsAngle(i), 0) & "," & Round(nearestNeighborsSortAngle(i), 0) & ")"
        'Next i
        
        
        ' Begin - Sort nearest neighbors by their angle representing the greatest right hand turn
        For i = 0 To nearestNeighborsCount - 1
            For j = i + 1 To nearestNeighborsCount - 1
                Dim tmpAngle As Double
                'cmpResult = PointOrientationSortCompare(points(points_oIdx(curPointIdx)), points(points_oIdx(nearestNeighborsIdx(i))), points(points_oIdx(nearestNeighborsIdx(j))))
                
                'cmpResult = IIf(nearestNeighborsSortAngle(j) < nearestNeighborsSortAngle(i), -1, 1)
                

                If nearestNeighborsSortAngle(j) > nearestNeighborsSortAngle(i) Then
                    tmpIdx = nearestNeighborsIdx(j)
                    nearestNeighborsIdx(j) = nearestNeighborsIdx(i)
                    nearestNeighborsIdx(i) = tmpIdx
                    
                    tmpAngle = nearestNeighborsAngle(j)
                    nearestNeighborsAngle(j) = nearestNeighborsAngle(i)
                    nearestNeighborsAngle(i) = tmpAngle
                    
                    tmpAngle = nearestNeighborsSortAngle(j)
                    nearestNeighborsSortAngle(j) = nearestNeighborsSortAngle(i)
                    nearestNeighborsSortAngle(i) = tmpAngle
                    
                    ' Note: We did not swap nearestNeighborsDistSq(i <-> j) since we don't need it anymore
                End If
            Next j
        Next i
        ' End - Sort nearest neighbors by their angle
        
        
        ' Begin - Find first candidate that does not have a path which crosses any of the hull edges
        Dim intersectsPrevious As Boolean
        
        i = 0
        nearestNeighborNextPointIdx = -1 ' -1 <-> undefined
        
        For i = 0 To nearestNeighborsCount - 1
            intersectsPrevious = False
            
            'Debug.Print "- " & siteNamesByPoint(points_oIdx(nearestNeighborsIdx(i)))
            
            For j = numPointsInHull - 1 To 1 Step -1 ' Note: start at end because we will most likely intersect with a closer edge
                
                'Debug.Print "  -- Test Its: " & _
                        siteNamesByPoint(points_oIdx(curPointIdx)) & ":" & siteNamesByPoint(points_oIdx(nearestNeighborsIdx(i))) & " and " & _
                        siteNamesByPoint(points_oIdx(hullPointsIdx(j))) & ":" & siteNamesByPoint(points_oIdx(hullPointsIdx(j - 1)))
                
                intersectsPrevious = LineSegmentsIntersectionByPoint( _
                    points(curPointIdx), points(nearestNeighborsIdx(i)), _
                    points(hullPointsIdx(j - 1)), points(hullPointsIdx(j)), _
                    False _
                )
                
                If intersectsPrevious = True Then
                    'Debug.Print "   -- Intersects"
                End If
                
                If intersectsPrevious = True Then Exit For
            Next j
            
            If intersectsPrevious = False Then
                nearestNeighborNextPointIdx = i
                Exit For
            End If
        Next i
        
        If nearestNeighborNextPointIdx = -1 Then
            ' We could not find a viable candidate in all the neighbors. Start over with a higher count for max neighbors
            Debug.Print "START OVER: " & "No viable candidate. maxNeighbors too small"
            maxNeighbors = maxNeighbors + 1
            GoTo StartOver
        End If
        ' End - Find first candidate which does not cross any of the hull edges
        
        'Debug.Print "- FOUND: " & siteNamesByPoint(points_oIdx(nextPointIdx))
        
        nextPointIdx = nearestNeighborsIdx(nearestNeighborNextPointIdx)
        prevAngle = nearestNeighborsAngle(nearestNeighborNextPointIdx)
        
        numPointsInHull = numPointsInHull + 1
        points_inHull(nextPointIdx) = True
        'points_hullOrder(nextPointIdx) = numPointsInHull
        hullPointsIdx(numPointsInHull - 1) = nextPointIdx
        
        prevPointIdx = curPointIdx
        curPointIdx = nextPointIdx
    Loop
    
    
    For i = 0 To numPointsInHull - 1
        Debug.Print points(hullPointsIdx(i)).ID
    Next i
    
    ' Create Hull as Polygon_2D
    HullPolygon.CoordinateCount = numPointsInHull
    
    ReDim HullPolygon.Coordinates(0 To numPointsInHull - 1)
    
    For i = 0 To numPointsInHull - 1
        HullPolygon.Coordinates(i) = points(hullPointsIdx(i))
    Next i
    
    HullPolygon.BoundingBox = Points_CalcBoundingRect(HullPolygon.Coordinates)
    
    ' Begin - Check if we missed any points, if so, then start over using more neighbors
    Dim allPointsInside As Boolean
    
    allPointsInside = True
    
    For i = 0 To numPoints - 1
        If points_inHull(i) = False Then ' Test non hull points only
            If PointInPolygon(points(i).x, points(i).y, HullPolygon) = False Then
                Debug.Print points(i).ID & " is not in hull polygon"
                
                allPointsInside = False
                Exit For
            End If
        End If
    Next i
    
    
    If allPointsInside = False Then
        ' We could not find a viable candidate in all the neighbors. Start over with a higher count for max neighbors
        Debug.Print "START OVER: " & "All points not bounded by hull. maxNeighbors too small (" & maxNeighbors & ")"
        maxNeighbors = maxNeighbors + 1
        GoTo StartOver
    End If
    ' End - Check if we missed any points, if so, then start over using more neighbors
    
    
    ConcaveHull = HullPolygon
    
    'For i = 0 To numPointsInHull - 1
    '    Debug.Print siteNamesByPoint(points_oIdx(hullPointsIdx(i)))
    'Next i
    
    
    
End Function

' ConvexHull: Given a set of points, finds the convex polygon that encloses all the points
'  - Uses Graham Scan algorithm. Also returns array of indices for the hull points in convexHullPointByIdx
'
Public Function ConvexHull(points() As Point_2D, Optional ByRef hullPointsByIdx As Variant) As Polygon_2D
    
    'Dim data() As Variant
    'data = Worksheet_LoadColumnDataToArray(ThisWorkbook.Sheets("ConvexHullPoints").Cells, Array("SITE_NAME", "LATITUDE", "LONGITUDE"))
    
    
   ' Dim points() As Point_2D
   ' Dim siteNamesByPoint() As String
   '
   ' Dim dpL As Long, dpU As Long, dpC As Long, dpI As Long
   ' dpL = LBound(data, 2)
   ' dpU = UBound(data, 2)
   ' dpC = dpU - dpL + 1
   '
   ' ReDim points(0 To dpC - 1) As Point_2D
   ' ReDim siteNamesByPoint(0 To dpC - 1) As String
   '
    'For dpI = 0 To dpC - 1
    '    siteNamesByPoint(dpI) = data(0, dpI + dpL)
    '    points(dpI).x = data(2, dpI + dpL)
    '    points(dpI).y = data(1, dpI + dpL)
    'Next dpI
        
    ' ---------------------------------------------------
    
    Dim i As Long, j As Long, n As Long, m As Long
    Dim tmpIdx As Long
    
    
    ' pointsIdx() will hold the the index of the original points so we don't have to perform swap/move operations on the points() array itself
    Dim pointsIdx() As Long
    n = UBound(points) - LBound(points) + 1
    
    ' Initial: not sorted:
    ReDim pointsIdx(0 To n - 1) As Long
    
    For i = 0 To n - 1: pointsIdx(i) = LBound(points) + i: Next i
  
    Dim ymin_idx As Long
    
    ymin_idx = 0
    
    For i = 0 To n - 1
        If ((points(pointsIdx(i)).y < points(pointsIdx(ymin_idx)).y) Or (points(pointsIdx(ymin_idx)).y = points(pointsIdx(i)).y And points(pointsIdx(i)).x < points(pointsIdx(ymin_idx)).x)) Then
            ymin_idx = i
        End If
    Next i

    
    
    ' Place bottom-most point at first position
    tmpIdx = pointsIdx(0)
    pointsIdx(0) = pointsIdx(ymin_idx)
    pointsIdx(ymin_idx) = tmpIdx

    ' Sort n-1 points with respect to the first point. A point p1 comes before p2 in sorted ouput
    ' if p2 as larger polar angle (in counterclockwise direction) than p1.
    For i = 1 To n - 1
        For j = i + 1 To n - 1
            Dim cmpResult As Integer
            cmpResult = PointOrientationSortCompare(points(pointsIdx(0)), points(pointsIdx(i)), points(pointsIdx(j)))
            
            'Debug.Print siteNamesByPoint(pointsIdx(0)) & "->" & siteNamesByPoint(pointsIdx(i)) & "->" & siteNamesByPoint(pointsIdx(j)) & ": " & IIf(cmpResult = -1, "CCW/Farther", "CC/Closer")
            
            If cmpResult > 0 Then
                tmpIdx = pointsIdx(j)
                pointsIdx(j) = pointsIdx(i)
                pointsIdx(i) = tmpIdx
            End If
        Next j
    Next i
    
    ' If two or more points make same angle with p0,
    ' Remove all but the one that is farthest from p0
    ' Remember that, in above sorting, our criteria was
    ' to keep the farthest point at the end when more than
    ' one points have same angle.
    m = 1 ' Initialize size of modified array
    For i = 1 To n - 1
        ' Keep removing i while angle of i and i+1 is same
        ' with respect to p0
        If i < n - 1 Then
             Do While i < n - 1 And OrientationByPoint(points(pointsIdx(0)), points(pointsIdx(i)), points(pointsIdx(i + 1))) = 0
                'Debug.Print "Same angle: " & siteNamesByPoint(pointsIdx(i)), siteNamesByPoint(pointsIdx(i + 1))
                i = i + 1
             Loop
        End If
    
        pointsIdx(m) = pointsIdx(i)
        m = m + 1 ' Update size of modified array
    Next i
    
    
    If m < 3 Then Err.Raise -1, , "Convex hull not possible"
    
    
    ' Not a real stack implementation, but we know the stack cannot grow to be more than m elements so we pre-allocate.
    Dim stack() As Long, stackTopIdx As Long
    ReDim stack(0 To m - 1) As Long
    
    ' Push first three points to stack
    stack(0) = pointsIdx(0)
    stack(1) = pointsIdx(1)
    stack(2) = pointsIdx(2)
    stackTopIdx = 2
    
    For i = 3 To m - 1
        Dim o As Integer

        ' While the orientation between the last two points on the stack and the current point, pop the stack until the orientation is CCW (non-left turn)
        Do While OrientationByPoint(points(stack(stackTopIdx - 1)), points(stack(stackTopIdx)), points(pointsIdx(i))) <> 2
            stackTopIdx = stackTopIdx - 1 ' Pops the stack
        Loop
    
        ' Push point to stack
        stackTopIdx = stackTopIdx + 1
        stack(stackTopIdx) = pointsIdx(i)
    Next i
    
    Dim ret As Polygon_2D
    
    hullPointsByIdx = stack
    
    ret.CoordinateCount = stackTopIdx + 1
    
    ReDim ret.Coordinates(0 To stackTopIdx) As Point_2D
    
    For i = 0 To stackTopIdx
        ret.Coordinates(i) = points(stack(i))
    Next i
    
    ret.BoundingBox = Points_CalcBoundingRect(ret.Coordinates)
    
    ConvexHull = ret

End Function
' Returns distance (in miles) and angle between two lat/long pairs.
'
' Uses a equirectangular plane approximation to the surface of a sphere near the given lat/lon. As a result, is it
' a fast computation but is valid for only lat/long pairs which are close to each other (<15mi apart)
Private Function EquirectangularDistanceVector(lat1 As Double, lon1 As Double, lat2 As Double, lon2 As Double) As Vector_2D


    Dim lat1_rad As Double
    Dim lon1_rad As Double
    Dim lat2_rad As Double
    Dim lon2_rad As Double
    Dim dlat As Double
    Dim dlon As Double
    Dim dist As Double

    Dim dx As Double, dy As Double

    Dim ret As Vector_2D

    ' Decimal degrees to radians
    lat1_rad = lat1 * PI / 180
    lon1_rad = lon1 * PI / 180
    lat2_rad = lat2 * PI / 180
    lon2_rad = lon2 * PI / 180

    dlat = lat2_rad - lat1_rad
    dlon = lon2_rad - lon1_rad

    ' calculate change in cartesian coordinates (units: radians)
    dx = dlon * Cos((lat1_rad + lat2_rad) / 2)
    dy = dlat


    ret.x = dx
    ret.y = dy

    ret.angle = (Atan2(dx, dy) * (180 / PI))
    ret.length = 0

    'If dx > EPSILON Or dy > EPSILON Then
    ret.length = Sqr(dx * dx + dy * dy) ' distance in radians
    ret.length = ret.length * EARTH_RADIUS_MI ' convert radians -> miles
    'End If

    EquirectangularDistanceVector = ret

End Function
Public Function EquirectangularDistance(lat1 As Double, lon1 As Double, lat2 As Double, lon2 As Double) As Double

    EquirectangularDistance = EquirectangularDistanceVector(lat1, lon1, lat2, lon2).length

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

Public Function Load_KML_Polygons_Old(kmlFile As String, Optional restrictToFolder As String = "") As Polygon_2D()

    ' 2015-04-17 - Now uses late-bound XML Dom objects
    
    
    
    Dim i As Long, j As Long
    
    
    Const KML_FOLDER_PATH_EXT_DELIMITER = "/"
    Const KML_FOLDER_PATH_INT_DELIMITER = "|||"
    
    Dim kmlFolderQ_Paths() As String
    Dim kmlFolderQ_Nodes() As Object 'MSXML2.IXMLDOMNode
    Dim kmlFolderQ_Count As Long
    Dim kmlFolderQ_ProcessedCount As Long
    
    
    Dim polygonPaths() As String
    Dim polygonCoordStr() As String
    Dim polygonCount As Long
    

    
    Dim xmlDoc As Object 'MSXML2.DOMDocument
    Dim xmlDocRoot As Object 'MSXML2.IXMLDOMNode
    
    Dim xmlNodePlacemarkList As Object 'MSXML2.IXMLDOMNodeList
    Dim xmlNodePlacemark As Object 'MSXML2.IXMLDOMNode
    
    Dim xmlNodePolygonList  As Object 'MSXML2.IXMLDOMNodeList
    Dim xmlNodePolygon As Object 'MSXML2.IXMLDOMNode
    
    
    Dim xmlNodePolygonCoordList  As Object 'MSXML2.IXMLDOMNodeList
    Dim xmlNodePolygonCoord As Object 'MSXML2.IXMLDOMNode
    
    Dim xmlNodeFolder  As Object 'MSXML2.IXMLDOMNode
    Dim xmlNodeFolderList  As Object 'MSXML2.IXMLDOMNodeList
    
    Dim xmlNodeName As Object 'MSXML2.IXMLDOMNode
    
    Dim xmlNode As Object 'MSXML2.IXMLDOMNode
    Dim xmlNodeList As Object 'MSXML2.IXMLDOMNodeList
    
    Dim xmlNodeTest As Object 'MSXML2.IXMLDOMNode



    ' ----------------------------------------------------------------------------------------------------------------
    ' Begin - Load KML and polygons coordinates
    ' ----------------------------------------------------------------------------------------------------------------
    Set xmlDoc = New MSXML2.DOMDocument
    'Set xmlDoc = CreateObject("MSXML2.DOMDocument")  'New MSXML2.DOMDocument
   
    If xmlDoc.Load(kmlFile) = False Then
        Err.Raise xmlDoc.parseError.ErrorCode, , xmlDoc.parseError.reason
    End If
    
   
    Set xmlDocRoot = xmlDoc.SelectSingleNode("/kml/Document")
    
    ReDim kmlFolderQ_Paths(1 To 1) As String
    ReDim kmlFolderQ_Node(1 To 1) As MSXML2.IXMLDOMNode
    ReDim kmlFolderQ_Processed(1 To 1) As Boolean
    
    kmlFolderQ_Count = 1
    kmlFolderQ_ProcessedCount = 0
    
    kmlFolderQ_Paths(1) = ""
    Set kmlFolderQ_Node(1) = xmlDocRoot
    kmlFolderQ_Processed(1) = False
    

    Dim idx As Long
    Dim nextIdx As Long
    
    Dim folderPathInternal As String, folderPathExternal As String
    
    Do While kmlFolderQ_ProcessedCount < kmlFolderQ_Count
        ' Get IDX of next folder Q item which has not been processed
        idx = 0
    
        For i = 1 To kmlFolderQ_Count
            If kmlFolderQ_Processed(i) = False Then
                idx = i
                Exit For
            End If
        Next i
        
        If idx = 0 Then Exit Do ' none found - we're done
        
        Set xmlNodeFolderList = kmlFolderQ_Node(idx).SelectNodes("Folder")
        
        If xmlNodeFolderList.length > 0 Then
            'nextIdx = kmlFolderQ_Count + 1
            'kmlFolderQ_Count = kmlFolderQ_Count + xmlNodeList.Length
        
            'ReDim Preserve kmlFolderQ_Paths(1 To kmlFolderQ_Count) As String
            'ReDim Preserve kmlFolderQ_Node(1 To kmlFolderQ_Count) As MSXML2.IXMLDOMNode
            'ReDim Preserve kmlFolderQ_Processed(1 To kmlFolderQ_Count) As Boolean
    
            For Each xmlNodeFolder In xmlNodeFolderList
                folderPathInternal = kmlFolderQ_Paths(idx) & xmlNodeFolder.SelectSingleNode("name").Text & KML_FOLDER_PATH_INT_DELIMITER
                folderPathExternal = Replace(folderPathInternal, KML_FOLDER_PATH_INT_DELIMITER, KML_FOLDER_PATH_EXT_DELIMITER)
                
                ' Check if folderpath is part of or within restrict folder path parameter
                Dim isFolderPathPartOfRestrictToFolderPath As Boolean ' determines if we continue traversing the folder
                Dim isFolderPathWithinRestrictToFolderPath As Boolean ' determines if we add polygons from this folder
                
                isFolderPathPartOfRestrictToFolderPath = True
                isFolderPathWithinRestrictToFolderPath = True
                
                If Len(restrictToFolder) > 0 Then
                    Dim folderPathCompareLength As Long
                    
                    folderPathCompareLength = IIf(Len(folderPathExternal) < Len(restrictToFolder), Len(folderPathExternal), Len(restrictToFolder))
                    
                    If Left(restrictToFolder, folderPathCompareLength) <> Left(folderPathExternal, folderPathCompareLength) Then isFolderPathPartOfRestrictToFolderPath = False
                
                
                    isFolderPathWithinRestrictToFolderPath = False
                
                    If isFolderPathPartOfRestrictToFolderPath = True And Len(folderPathExternal) >= Len(restrictToFolder) Then
                        If restrictToFolder = Left(folderPathExternal, Len(restrictToFolder)) Then isFolderPathWithinRestrictToFolderPath = True
                    End If
                End If
                
                'Debug.Print folderPathExternal, isFolderPathPartOfrestrictToFolderPath, isFolderPathWithinrestrictToFolderPath
                
                ' Folder path is part of the restrict folder path -> contiue traversing subfolders
                If isFolderPathPartOfRestrictToFolderPath Then
                    kmlFolderQ_Count = kmlFolderQ_Count + 1
                    nextIdx = kmlFolderQ_Count
                    
                    ReDim Preserve kmlFolderQ_Paths(1 To kmlFolderQ_Count) As String
                    ReDim Preserve kmlFolderQ_Node(1 To kmlFolderQ_Count) As MSXML2.IXMLDOMNode
                    ReDim Preserve kmlFolderQ_Processed(1 To kmlFolderQ_Count) As Boolean
    
                    Set kmlFolderQ_Node(nextIdx) = xmlNodeFolder
                    kmlFolderQ_Paths(nextIdx) = folderPathInternal
                    kmlFolderQ_Processed(nextIdx) = False
                End If
                
                ' Folder path is part of the restrict folder path -> load polygons
                If isFolderPathWithinRestrictToFolderPath Then
                    ' Get Placemarks nodes which define polygons
                    Set xmlNodePlacemarkList = xmlNodeFolder.SelectNodes("./Placemark[Polygon] | ./Placemark[MultiGeometry/Polygon]")
                    
                                    
                    For Each xmlNodePlacemark In xmlNodePlacemarkList
                        'Set xmlNodeName = xmlNodePlacemark.SelectSingleNode("name")
                        'Debug.Print xmlNodePlacemark.SelectSingleNode("name").Text
                        
                        
                        
                        Set xmlNodePolygonCoordList = xmlNodePlacemark.SelectNodes("./Polygon[outerBoundaryIs/LinearRing/coordinates] | ./MultiGeometry/Polygon[outerBoundaryIs/LinearRing/coordinates]")
                        
                        
                        'Debug.Print "Adding Polygon: "; folderPathExternal & xmlNodePlacemark.SelectSingleNode("name").Text & " (shapes: " & xmlNodePolygonCoordList.Length & ")"
                        
                        For Each xmlNodePolygonCoord In xmlNodePolygonCoordList
                            
                            polygonCount = polygonCount + 1
                            
                            ReDim Preserve polygonPaths(1 To polygonCount) As String
                            ReDim Preserve polygonCoordStr(1 To polygonCount) As String
                        
                            polygonPaths(polygonCount) = folderPathExternal & xmlNodePlacemark.SelectSingleNode("name").Text
                            polygonCoordStr(polygonCount) = xmlNodePolygonCoord.Text
                        Next xmlNodePolygonCoord
                    
                    Next xmlNodePlacemark
    
                    Set xmlNodePlacemarkList = Nothing
                    Set xmlNodePlacemark = Nothing
                    
    
                End If
                
                'nextIdx = nextIdx + 1
                    
            Next xmlNodeFolder
            
        End If
        
        Set xmlNodeFolder = Nothing
        Set xmlNodeFolderList = Nothing
        
        Set kmlFolderQ_Node(idx) = Nothing
        kmlFolderQ_Processed(idx) = True
    
        kmlFolderQ_ProcessedCount = kmlFolderQ_ProcessedCount + 1
    Loop


    Set xmlDoc = Nothing
    Set xmlNodeFolder = Nothing
    Set xmlNode = Nothing
    
    ' ----------------------------------------------------------------------------------------------------------------
    ' End - Load KML and polygons coordinates
    ' ----------------------------------------------------------------------------------------------------------------
    
    
    ' ----------------------------------------------------------------------------------------------------------------
    ' Begin - Convert polygon coordinates to Polygon_2D structures where X=Longitude,Y=Latitude
    ' ----------------------------------------------------------------------------------------------------------------
    Dim polygons() As Polygon_2D
    ReDim polygons(1 To polygonCount) As Polygon_2D
    
    Dim polygonCoordinates() As Point_2D
    Dim polygonCoordinatesActualCount As Long
    
    
    'Dim i As Long, j As Long
    
    Dim polygonCoordPairsArr() As String
    Dim polygonCoordsArr() As String

    
    Dim coordY As Double, coordX As Double
    
    For i = 1 To polygonCount
        polygons(i).ID = polygonPaths(i)
    
        polygonCoordPairsArr = Split(polygonCoordStr(i), " ")
        
        
        polygons(i).CoordinateCount = UBound(polygonCoordPairsArr) - LBound(polygonCoordPairsArr) + 1
        polygonCoordinatesActualCount = 0
        
        ReDim polygonCoordinates(1 To polygons(i).CoordinateCount) As Point_2D
        
        
        For j = LBound(polygonCoordPairsArr) To UBound(polygonCoordPairsArr)
            Dim tmp As String
            tmp = polygonCoordPairsArr(j)
            
            
            'Debug.Print polygonCoordPairsArr(j)
            
            If InStr(1, polygonCoordPairsArr(j), ",") > 0 Then
                polygonCoordsArr = Split(polygonCoordPairsArr(j), ",")
                
                coordX = CDbl(polygonCoordsArr(LBound(polygonCoordsArr) + 0))
                coordY = CDbl(polygonCoordsArr(LBound(polygonCoordsArr) + 1))
                
                
                polygonCoordinatesActualCount = polygonCoordinatesActualCount + 1
                
                polygonCoordinates(polygonCoordinatesActualCount).x = coordX
                polygonCoordinates(polygonCoordinatesActualCount).y = coordY
            End If
            
        Next j
        
        If polygonCoordinatesActualCount < polygons(i).CoordinateCount Then
            ReDim Preserve polygonCoordinates(1 To polygonCoordinatesActualCount) As Point_2D
        End If
        
        polygons(i).Coordinates = polygonCoordinates
        polygons(i).CoordinateCount = polygonCoordinatesActualCount
        
        polygons(i).BoundingBox = Points_CalcBoundingRect(polygons(i).Coordinates)
    
    Next i
    ' ----------------------------------------------------------------------------------------------------------------
    ' End - Convert polygon coordinates to Polygon_2D structures where X=Longitude,Y=Latitude
    ' ----------------------------------------------------------------------------------------------------------------
    
    
    Load_KML_Polygons_Old = polygons
    

End Function


Public Function Load_KML_Polygons(kmlFile As String, Optional restrictToPath As String = "") As Polygon_2D()

    ' 2015-04-17 - Now uses late-bound XML Dom objects
    
    
    
    Dim i As Long, j As Long
    
    
    Const KML_FOLDER_PATH_EXT_DELIMITER = "/"
    Const KML_FOLDER_PATH_INT_DELIMITER = "|||"
    
    Dim kmlFolderQ_Paths() As String
    Dim kmlFolderQ_Nodes() As Object 'MSXML2.IXMLDOMNode
    Dim kmlFolderQ_Count As Long
    Dim kmlFolderQ_ProcessedCount As Long
    
    
    Dim polygonPaths() As String
    Dim polygonCoordStr() As String
    Dim polygonCount As Long
    

    
    Dim xmlDoc As Object 'MSXML2.DOMDocument
    Dim xmlDocRoot As Object 'MSXML2.IXMLDOMNode
    
    Dim xmlNodePlacemarkList As Object 'MSXML2.IXMLDOMNodeList
    Dim xmlNodePlacemark As Object 'MSXML2.IXMLDOMNode
    
    Dim xmlNodePolygonList  As Object 'MSXML2.IXMLDOMNodeList
    Dim xmlNodePolygon As Object 'MSXML2.IXMLDOMNode
    
    
    Dim xmlNodePolygonCoordList  As Object 'MSXML2.IXMLDOMNodeList
    Dim xmlNodePolygonCoord As Object 'MSXML2.IXMLDOMNode
    
    Dim xmlNodeFolder  As Object 'MSXML2.IXMLDOMNode
    Dim xmlNodeFolderList  As Object 'MSXML2.IXMLDOMNodeList
    
    Dim xmlNodeName As Object 'MSXML2.IXMLDOMNode
    
    Dim xmlNode As Object 'MSXML2.IXMLDOMNode
    Dim xmlNodeList As Object 'MSXML2.IXMLDOMNodeList
    
    Dim xmlNodeTest As Object 'MSXML2.IXMLDOMNode



    ' ----------------------------------------------------------------------------------------------------------------
    ' Begin - Load KML and polygons coordinates
    ' ----------------------------------------------------------------------------------------------------------------
    Set xmlDoc = New MSXML2.DOMDocument
    'Set xmlDoc = CreateObject("MSXML2.DOMDocument")  'New MSXML2.DOMDocument
   
    If xmlDoc.Load(kmlFile) = False Then
        Err.Raise xmlDoc.parseError.ErrorCode, , xmlDoc.parseError.reason
    End If
    
   
    Set xmlDocRoot = xmlDoc.SelectSingleNode("/kml/Document")
    
    ReDim kmlFolderQ_Paths(1 To 1) As String
    ReDim kmlFolderQ_Node(1 To 1) As MSXML2.IXMLDOMNode
    ReDim kmlFolderQ_Processed(1 To 1) As Boolean
    
    kmlFolderQ_Count = 1
    kmlFolderQ_ProcessedCount = 0
    
    kmlFolderQ_Paths(1) = ""
    Set kmlFolderQ_Node(1) = xmlDocRoot
    kmlFolderQ_Processed(1) = False
    

    Dim idx As Long
    Dim nextIdx As Long
    
    
    Dim isPathPartOfRestrictToPath As Boolean ' determines if we continue traversing the folder
    Dim isPathWithinRestrictToPath As Boolean ' determines if we add polygons from this folder
    
    Dim pathInternal As String, pathExternal As String
    Dim minLength As Long
    
    Do While kmlFolderQ_ProcessedCount < kmlFolderQ_Count
        ' Get IDX of next folder Q item which has not been processed
        idx = 0
    
        For i = 1 To kmlFolderQ_Count
            If kmlFolderQ_Processed(i) = False Then
                idx = i
                Exit For
            End If
        Next i
        
        If idx = 0 Then Exit Do ' none found - we're done
        
        
        
        ' --------------------------------------------------
        ' Begin - Add placemark polygons within current node
        ' --------------------------------------------------
        ' Get Placemarks nodes which define polygons
        Set xmlNodePlacemarkList = kmlFolderQ_Node(idx).SelectNodes("./Placemark[Polygon] | ./Placemark[MultiGeometry/Polygon]")
        
                        
        For Each xmlNodePlacemark In xmlNodePlacemarkList
            'Set xmlNodeName = xmlNodePlacemark.SelectSingleNode("name")
            pathInternal = kmlFolderQ_Paths(idx) & xmlNodePlacemark.SelectSingleNode("name").Text
            pathExternal = Replace(pathInternal, KML_FOLDER_PATH_INT_DELIMITER, KML_FOLDER_PATH_EXT_DELIMITER)
            
            
            'Debug.Print "Checking polygon: " & pathExternal
            
            
            isPathWithinRestrictToPath = True
            
            If Len(restrictToPath) > 0 Then
                If Len(restrictToPath) <= Len(pathExternal) Then
                    isPathWithinRestrictToPath = (StrComp(restrictToPath, Left(pathExternal, Len(restrictToPath)), vbTextCompare) = 0)
                Else
                    isPathWithinRestrictToPath = False
                End If
            End If

            If isPathWithinRestrictToPath Then
                Set xmlNodePolygonCoordList = xmlNodePlacemark.SelectNodes("./Polygon[outerBoundaryIs/LinearRing/coordinates] | ./MultiGeometry/Polygon[outerBoundaryIs/LinearRing/coordinates]")
                
                'Debug.Print "Adding Polygon: "; pathExternal & " (shapes: " & xmlNodePolygonCoordList.length & ")"
                
                For Each xmlNodePolygonCoord In xmlNodePolygonCoordList
                    polygonCount = polygonCount + 1
                    
                    ReDim Preserve polygonPaths(1 To polygonCount) As String
                    ReDim Preserve polygonCoordStr(1 To polygonCount) As String
                
                    polygonPaths(polygonCount) = pathExternal
                    polygonCoordStr(polygonCount) = xmlNodePolygonCoord.Text
                Next xmlNodePolygonCoord
            End If
        
        Next xmlNodePlacemark

        Set xmlNodePlacemarkList = Nothing
        Set xmlNodePlacemark = Nothing
        ' --------------------------------------------------
        ' End - Add placemark polygons within current node
        ' --------------------------------------------------
                
        
        
        ' --------------------------------------------------
        ' Begin - Add subfolders to queue
        ' --------------------------------------------------
        Set xmlNodeFolderList = kmlFolderQ_Node(idx).SelectNodes("Folder")
        
        If xmlNodeFolderList.length > 0 Then
            'nextIdx = kmlFolderQ_Count + 1
            'kmlFolderQ_Count = kmlFolderQ_Count + xmlNodeList.Length
        
            'ReDim Preserve kmlFolderQ_Paths(1 To kmlFolderQ_Count) As String
            'ReDim Preserve kmlFolderQ_Node(1 To kmlFolderQ_Count) As MSXML2.IXMLDOMNode
            'ReDim Preserve kmlFolderQ_Processed(1 To kmlFolderQ_Count) As Boolean
    
            For Each xmlNodeFolder In xmlNodeFolderList
                pathInternal = kmlFolderQ_Paths(idx) & xmlNodeFolder.SelectSingleNode("name").Text & KML_FOLDER_PATH_INT_DELIMITER
                pathExternal = Replace(pathInternal, KML_FOLDER_PATH_INT_DELIMITER, KML_FOLDER_PATH_EXT_DELIMITER)
                
                ' Check if folderpath is part of or within restrict folder path parameter
                
                isPathPartOfRestrictToPath = True
                isPathWithinRestrictToPath = True
                
                If Len(restrictToPath) > 0 Then
                    minLength = IIf(Len(pathExternal) < Len(restrictToPath), Len(pathExternal), Len(restrictToPath))
                    
                    If StrComp(Left(restrictToPath, minLength), Left(pathExternal, minLength), vbTextCompare) <> 0 Then isPathPartOfRestrictToPath = False
                
                
                    'isPathWithinRestrictToPath = False
                
                    'If isPathPartOfRestrictToPath = True And Len(pathExternal) >= Len(restrictToPath) Then
                    '    If restrictToPath = Left(pathExternal, Len(restrictToPath)) Then isPathWithinRestrictToPath = True
                    'End If
                End If
                
                'Debug.Print pathExternal, isPathPartOfRestrictToPath, isPathWithinRestrictToPath
                
                ' Folder path is part of the restrict folder path -> contiue traversing subfolders
                If isPathPartOfRestrictToPath Then
                    kmlFolderQ_Count = kmlFolderQ_Count + 1
                    nextIdx = kmlFolderQ_Count
                    
                    ReDim Preserve kmlFolderQ_Paths(1 To kmlFolderQ_Count) As String
                    ReDim Preserve kmlFolderQ_Node(1 To kmlFolderQ_Count) As MSXML2.IXMLDOMNode
                    ReDim Preserve kmlFolderQ_Processed(1 To kmlFolderQ_Count) As Boolean
    
                    Set kmlFolderQ_Node(nextIdx) = xmlNodeFolder
                    kmlFolderQ_Paths(nextIdx) = pathInternal
                    kmlFolderQ_Processed(nextIdx) = False
                End If
                
                ' Folder path is part of the restrict folder path -> load polygons
                'nextIdx = nextIdx + 1
                    
            Next xmlNodeFolder
            
        End If
        
        Set xmlNodeFolder = Nothing
        Set xmlNodeFolderList = Nothing
        ' --------------------------------------------------
        ' End - Add subfolders to queue
        ' --------------------------------------------------
        
        Set kmlFolderQ_Node(idx) = Nothing
        kmlFolderQ_Processed(idx) = True
    
        kmlFolderQ_ProcessedCount = kmlFolderQ_ProcessedCount + 1
    Loop


    Set xmlDoc = Nothing
    Set xmlNodeFolder = Nothing
    Set xmlNode = Nothing
    
    ' ----------------------------------------------------------------------------------------------------------------
    ' End - Load KML and polygons coordinates
    ' ----------------------------------------------------------------------------------------------------------------
    
    If polygonCount = 0 Then Exit Function
    
    ' ----------------------------------------------------------------------------------------------------------------
    ' Begin - Convert polygon coordinates to Polygon_2D structures where X=Longitude,Y=Latitude
    ' ----------------------------------------------------------------------------------------------------------------
    Dim polygons() As Polygon_2D
    ReDim polygons(1 To polygonCount) As Polygon_2D
    
    Dim polygonCoordinates() As Point_2D
    Dim polygonCoordinatesActualCount As Long
    
    
    'Dim i As Long, j As Long
    
    Dim polygonCoordPairsArr() As String
    Dim polygonCoordsArr() As String

    
    Dim coordY As Double, coordX As Double
    
    For i = 1 To polygonCount
        polygons(i).ID = polygonPaths(i)
    
        polygonCoordPairsArr = Split(polygonCoordStr(i), " ")
        
        
        polygons(i).CoordinateCount = UBound(polygonCoordPairsArr) - LBound(polygonCoordPairsArr) + 1
        polygonCoordinatesActualCount = 0
        
        ReDim polygonCoordinates(1 To polygons(i).CoordinateCount) As Point_2D
        
        
        For j = LBound(polygonCoordPairsArr) To UBound(polygonCoordPairsArr)
            Dim tmp As String
            tmp = polygonCoordPairsArr(j)
            
            
            'Debug.Print polygonCoordPairsArr(j)
            
            If InStr(1, polygonCoordPairsArr(j), ",") > 0 Then
                polygonCoordsArr = Split(polygonCoordPairsArr(j), ",")
                
                coordX = CDbl(polygonCoordsArr(LBound(polygonCoordsArr) + 0))
                coordY = CDbl(polygonCoordsArr(LBound(polygonCoordsArr) + 1))
                
                
                polygonCoordinatesActualCount = polygonCoordinatesActualCount + 1
                
                polygonCoordinates(polygonCoordinatesActualCount).x = coordX
                polygonCoordinates(polygonCoordinatesActualCount).y = coordY
            End If
            
        Next j
        
        If polygonCoordinatesActualCount < polygons(i).CoordinateCount Then
            ReDim Preserve polygonCoordinates(1 To polygonCoordinatesActualCount) As Point_2D
        End If
        
        polygons(i).Coordinates = polygonCoordinates
        polygons(i).CoordinateCount = polygonCoordinatesActualCount
        
        polygons(i).BoundingBox = Points_CalcBoundingRect(polygons(i).Coordinates)
    
    Next i
    ' ----------------------------------------------------------------------------------------------------------------
    ' End - Convert polygon coordinates to Polygon_2D structures where X=Longitude,Y=Latitude
    ' ----------------------------------------------------------------------------------------------------------------
    
    
    Load_KML_Polygons = polygons
    

End Function

Public Function KML_Polygon_HitTest(kmlFile As String, srcLats As Variant, srcLons As Variant, Optional restrictToFolder As String = "", Optional shortPolygonIDs = True) As Variant()

    
    ' Requires a reference to Microsoft XML, v6.0 (Tools -> References -> Check Microsoft XML)


    'Dim kmlFile As String: kmlFile = "C:\Users\w503686\Documents\Google Earth\LicensesSpectrum\doc.kml"
    'Dim srcLats As Variant: Set srcLats = Range("Sheet1!B1:B12")
    'Dim srcLons As Variant: Set srcLons = Range("Sheet1!C1:C12")
    'Dim restrictToFolder As String: restrictToFolder = "Licenses/Spectrum" '"Licenses/Spectrum/AWS"
    'Dim shortPolygonIDs As Boolean: shortPolygonIDs = True
    
    
    
    
    Dim i As Long, j As Long, k As Long
    Dim rngLats As Range, rngLons As Range, rngData() As Variant
    
    
    Dim pointsLat() As Double, pointsLon() As Double
    Dim pointsIsValid() As Boolean
    Dim pointsPolygons() As String
    Dim pointsCount As Long
    
    Dim tn As String
    
    tn = TypeName(srcLats)
    
    If TypeName(srcLats) = "Range" Then
        Set rngLats = srcLats
        Set rngLons = srcLons
        
        If rngLats.rows.count <> rngLons.rows.count Then
            KML_Polygon_HitTest = Array("Invalid set of latitude/longitude pairs")
            Exit Function
        End If
        
        pointsCount = rngLats.rows.count
        
        ReDim pointsLat(1 To pointsCount) As Double
        ReDim pointsLon(1 To pointsCount) As Double
        ReDim pointsIsValid(1 To pointsCount) As Boolean
    
        On Error Resume Next
        For i = 1 To pointsCount
            pointsLat(i) = CDbl(rngLats.Cells(i, 1).Value)
            pointsLon(i) = CDbl(rngLons.Cells(i, 1).Value)
            pointsIsValid(i) = IIf(pointsLat(i) <> 0 And pointsLon(i) <> 0, True, False)
        Next i
        On Error GoTo 0
    ElseIf IsNumeric(srcLats) Then
        pointsCount = 1
        ReDim pointsLat(1 To 1) As Double
        ReDim pointsLon(1 To 1) As Double
        ReDim pointsIsValid(1 To pointsCount) As Boolean
        pointsLat(1) = srcLats
        pointsLon(1) = srcLons
        pointsIsValid(1) = IIf(pointsLat(1) <> 0 And pointsLon(1) <> 0, True, False)
    End If
    
    If pointsCount = 0 Then
        KML_Polygon_HitTest = Array("Invalid set of latitude/longitude pairs")
        Exit Function
    End If
    
    
    ReDim pointsPolygons(1 To pointsCount) As String
    
    
    Dim polygons() As Polygon_2D
    Dim sX As Double, sY As Double
    
    polygons = Load_KML_Polygons(kmlFile, restrictToFolder)
    
    ' If shortPolygonIDs is TRUE, remove the beginning of the path (aka. restrictToFolder path)
    If shortPolygonIDs And Len(restrictToFolder) > 0 Then
        For j = LBound(polygons) To UBound(polygons)
            polygons(j).ID = Right(polygons(j).ID, Len(polygons(j).ID) - Len(restrictToFolder) - 1)
        Next j
    End If
    
    
    Const POLYGON_DELIM = ";"
    
    For i = 1 To pointsCount
        If pointsIsValid(i) Then
            sY = pointsLat(i)
            sX = pointsLon(i)
            
            pointsPolygons(i) = ""
            
            For j = LBound(polygons) To UBound(polygons)
                If PointInPolygon(sX, sY, polygons(j)) Then
                    pointsPolygons(i) = pointsPolygons(i) & polygons(j).ID & POLYGON_DELIM
                End If
            Next j
            
            If Len(pointsPolygons(i)) > Len(POLYGON_DELIM) Then pointsPolygons(i) = Left(pointsPolygons(i), Len(pointsPolygons(i)) - Len(POLYGON_DELIM))
        End If
    Next i
    
    
    Dim retArr() As Variant
    
    Dim pointsPolygonsArr() As String
    Dim pointsPolygonIdx As Long
    
        
        
    ' If calling directly from excel as an array formula, then modify the result to fit the cells
    If IsObject(Application.Caller) Then
        Dim callerRows As Long, callerCols As Long
        Dim isLastCellInCallerArray As Boolean
    
        callerRows = Application.Caller.rows.count
        callerCols = Application.Caller.columns.count
        
        
        ReDim retArr(1 To callerRows, 1 To callerCols) As Variant
            
        
        For i = 1 To callerRows
            If i <= pointsCount Then
            
                If pointsIsValid(i) = False Then
                    For j = 1 To callerCols
                        retArr(i, j) = CVErr(XlCVError.xlErrValue)
                    Next j
                ElseIf pointsIsValid(i) = True And Len(pointsPolygons(i)) > 0 Then
                        pointsPolygonsArr = Split(pointsPolygons(i), POLYGON_DELIM)
                        pointsPolygonIdx = LBound(pointsPolygonsArr)
            
                        For j = 1 To callerCols
                            isLastCellInCallerArray = IIf(j = callerCols, True, False)
                            
                            If isLastCellInCallerArray And pointsPolygonIdx < UBound(pointsPolygonsArr) Then
                                ' Caller array column size is less than number of polygons
                                ' Fill the last array item with the rest of the list seperated by the delimiter
                                retArr(i, j) = pointsPolygonsArr(pointsPolygonIdx)
                                
                                For k = pointsPolygonIdx + 1 To UBound(pointsPolygonsArr)
                                    retArr(i, j) = retArr(i, j) & POLYGON_DELIM & pointsPolygonsArr(k)
                                Next k
                            Else
                                If pointsPolygonIdx > UBound(pointsPolygonsArr) Then
                                    ' Caller array size is greater than expanded cluster size
                                    ' Fill array item with empty string
                                    retArr(i, j) = ""
                                Else
                                    retArr(i, j) = pointsPolygonsArr(pointsPolygonIdx)
                                    pointsPolygonIdx = pointsPolygonIdx + 1
                                End If
                            End If
                        Next j
                End If
                
            Else
            
                For j = 1 To callerCols
                    retArr(i, j) = CVErr(XlCVError.xlErrNA)
                Next j
            
            End If
            
        Next i
    Else
        ReDim retArr(1 To pointsCount) As Variant
        
        For i = 1 To pointsCount
            If pointsIsValid(i) Then
                retArr(i) = pointsPolygons(i)
            Else
                retArr(i) = CVErr(XlCVError.xlErrValue)
            End If
        Next i
    End If
    
    
    
    KML_Polygon_HitTest = retArr


End Function

Private Function PointInPolygon(x As Double, y As Double, polygon As Polygon_2D) As Boolean

' Referenced from: http://alienryderflex.com/polygon/ which is the implementation in C++

    ' Point is not in polygon bounding rectangle -> then no need to further investigate
    If Not (polygon.BoundingBox.TopLeft.y < y And y < polygon.BoundingBox.BottomRight.y And polygon.BoundingBox.TopLeft.x < x And x < polygon.BoundingBox.BottomRight.x) Then
        PointInPolygon = False
        Exit Function
    End If

        
    Dim i As Long, j As Long, polySides As Integer
    Dim oddNodes As Boolean
    
    oddNodes = False
    
    j = UBound(polygon.Coordinates)
    
    For i = LBound(polygon.Coordinates) To UBound(polygon.Coordinates)
    
        If (((polygon.Coordinates(i).y < y And polygon.Coordinates(j).y >= y) _
            Or (polygon.Coordinates(j).y < y And polygon.Coordinates(i).y >= y)) _
            And (polygon.Coordinates(i).x <= x Or polygon.Coordinates(j).x <= x)) Then
            
            oddNodes = oddNodes Xor (polygon.Coordinates(i).x + (y - polygon.Coordinates(i).y) / (polygon.Coordinates(j).y - polygon.Coordinates(i).y) * (polygon.Coordinates(j).x - polygon.Coordinates(i).x) < x)
        End If
        
        j = i
    Next i
    
    PointInPolygon = oddNodes

End Function

Public Function PlacemarkInArc(pmLat As Double, pmLon As Double, arcLat As Double, arcLon As Double, arcAzimuth As Double, arcRadius As Double, arcWidth As Double) As Boolean

    'Dim c As Variant, d As Variant
    
    'c = Array(40.058167, -75.238513)
    'd = Array(40.0622, -75.2393)

    'Dim pmLat As Double, pmLon As Double
    'Dim arcLat As Double, arcLon As Double
    'Dim arcAngle As Double, arcRadius As Double
    'Dim arcWidth As Double
    
    
    'arcLat = c(0): arcLon = c(1)
    'pmLat = d(0): pmLon = d(1)
    
    'arcAngle = Angle_Normalize180(90 - 10)
    'arcWidth = 67
    'arcRadius = 0.5
    
    Dim arcAngle As Double
    arcAngle = Angle_Normalize180(90 - arcAzimuth)
    
    Dim distVec As Vector_2D
    Dim angleDiff As Double
    
    distVec = EquirectangularDistanceVector(arcLat, arcLon, pmLat, pmLon)
    

    PlacemarkInArc = False
    
    ' Distance is greater than arcRadius? -> Not in ARC
    If distVec.length <= arcRadius + EPSILON Then
        angleDiff = Angle_SmallestDifference(distVec.angle, arcAngle)
        
        If angleDiff < arcWidth / 2 + EPSILON Then
            PlacemarkInArc = True
        End If
    End If
    

End Function


' Calculates line of sight characteristics between a lat/lon pair.
'
'   Uses Google Elevation API to query the elevation profile and determines. All returned measurements are in feet.
'
'   Returns:
'       HasLineOfSight - True if there isn't any terrain obstructing the view from point1 to point2
'       ElevationStart - The elevation at lat/lon pair 1
'       ElevationEnd - The elevation at lat/lon pair 2
'       ElevationGain - Difference between ElevationStart and ElevationEnd
'       MaxPathElevation - The highest elevation point in the elevation profile (except for end elevation)
'       MinLineOfSightElevation - Minimum required elevation for HasLineOfSight to be true (can be less than ElevationStart if ElevationGain > 0)
Public Function Geospatial_CalcLineOfSight(lat1 As Double, lon1 As Double, lat2 As Double, lon2 As Double, Optional samplesPerMile = 20) As Geospatial_LineOfSight_Result

    'Dim lat1 As Double, lon1 As Double, lat2 As Double, lon2 As Double
    'Dim samplesPerMile As Integer: samplesPerMile = 20
    
    
    'Dim c1 As Variant, c2 As Variant
    
    'c1 = Array(40.020111, -75.217958)
    'c2 = Array(40.0172, -75.2408)
    
    ''c1 = Array(40.253725, -74.861311)
    'c2 = Array(40.275664, -74.827664)
    
    'lat1 = c1(0): lon1 = c1(1)
    'lat2 = c2(0): lon2 = c2(1)
    
    
    
    Const FEET_PER_METER As Double = 3.28084
    
    Dim dist As Double, samples As Integer
    
    dist = EquirectangularDistance(lat1, lon1, lat2, lon2)
    
    samples = Round(samplesPerMile * dist + 0.5) ' the addition of +0.5 forces a round up
    If samples < 10 Then samples = 10 ' minumum of 10 samples
    
    
    
    Dim samplesPathResult As Variant
    'Dim elevationProfile() As Variant
    Dim elevationProfile() As Double
    
    'Dim retArr() As Double

    samplesPathResult = Google_Elevation_SampledPath_1D(lat1, lon1, lat2, lon2, samples)
    elevationProfile = samplesPathResult ' Implicitly converts Variant to Double()
    
    'elevationProfile = Array(13.8364964, 35.7976036, 52.0680275, 62.0587387, 50.473587, 47.463913, 55.7948074, 61.957737, 61.5434914, 64.019249, 64.2760696, 66.3294525, 75.3158875, 75.8281708, 79.1792679, 80.1346893, 78.3832092, 74.4405899, 65.016098, 76.3104019)
   

    Dim i As Long
    
    ' convert elevation profile to feet (google returns meters)
    For i = LBound(elevationProfile) To UBound(elevationProfile)
        elevationProfile(i) = elevationProfile(i) * FEET_PER_METER
    Next i
    
    
    Dim ret As Geospatial_LineOfSight_Result

    Dim startElevation As Double, endElevation As Double
    Dim elevationProfilePointCount As Long
    
    elevationProfilePointCount = Array_Count(elevationProfile)
    
    startElevation = elevationProfile(LBound(elevationProfile))
    endElevation = elevationProfile(UBound(elevationProfile))
    
    Dim maxElevationIdx As Long
    Dim los_slope As Double, los_elevation As Double
    Dim los_elevationAtMax As Double
    
    
    ' Calculation maximum elevation between start/end points
    maxElevationIdx = LBound(elevationProfile)  ' start with first sample
    
    For i = LBound(elevationProfile) To UBound(elevationProfile) - 1 ' do not include the ending elevation in the max
        If elevationProfile(i) > elevationProfile(maxElevationIdx) Then maxElevationIdx = i
    Next i
    
    
    ' LOS elevation is linear interpolation between start and ending elevations
    los_slope = (endElevation - startElevation) / (elevationProfilePointCount - 1)
    
    ' Calculation LOS elevation at maximum true elevation
    los_elevationAtMax = elevationProfile(maxElevationIdx)
    

    Dim minStartElev_slope As Double, minStartElev As Double, newLos As Double
    
    
    minStartElev_slope = (endElevation - los_elevationAtMax) / (UBound(elevationProfile) - maxElevationIdx)
    minStartElev = los_elevationAtMax - minStartElev_slope * (maxElevationIdx - LBound(elevationProfile))
    
    For i = LBound(elevationProfile) To UBound(elevationProfile)
        los_elevation = startElevation + los_slope * (i - LBound(elevationProfile))
        newLos = minStartElev + minStartElev_slope * (i - LBound(elevationProfile))
        Debug.Print Format(elevationProfile(i), "0.00"), Format(los_elevation, "0.00"), Format(newLos, "0.00"), elevationProfile(i) > los_elevation, IIf(i = maxElevationIdx, "*", "")
    Next i
    
    
    ret.ElevationStart = startElevation
    ret.ElevationEnd = endElevation
    ret.ElevationGain = endElevation - startElevation
    ret.MaxPathElevation = los_elevationAtMax
    ret.MinLineOfSightElevation = minStartElev
    
    ret.HasLineOfSight = IIf(minStartElev <= startElevation, True, False)
    
    
    Geospatial_CalcLineOfSight = ret
    
    
End Function
Public Function Google_GeoCode(address As String, Optional googleApiKey As String) As Google_GeoCode_Result

    Dim xmlHttpReq As Object 'New XMLHTTP30
    Dim xmlDocResults As Object 'DOMDocument30
    Dim xmlStatusNode As Object 'IXMLDOMNode


    Set xmlHttpReq = CreateObject("Microsoft.XMLHTTP")

    'On Error GoTo errorHandler


    If googleApiKey = "" Then googleApiKey = "AIzaSyDWrKMGCwpGG299G6rLMF5FnMuIZMugpoE"
    
    Dim url As String
    
    url = "https://maps.googleapis.com/maps/api/geocode/xml?" & URL_BuildQueryString( _
        "key", googleApiKey, _
        "address", address _
    )
    

    'Send the request to the Google server.
    xmlHttpReq.Open "GET", url, False
    xmlHttpReq.Send

    'Read the results from the request.
    Set xmlDocResults = xmlHttpReq.responseXML
    'Results.LoadXML Request.responseText

    'Get the status node value.
    Set xmlStatusNode = xmlDocResults.SelectSingleNode("//status")

    Dim ret As Google_GeoCode_Result
    
    
    ret.status_code = UCase(xmlStatusNode.Text)

    'Based on the status node result, proceed accordingly.
    Select Case ret.status_code
        Case "OK"   'The API request was successful. At least one geocode was returned.

            Dim xmlAddressComponentNode As Object 'IXMLDOMNode
            Dim xmlAddressComponentNodeList As Object 'IXMLDOMNodeList


            Set xmlAddressComponentNodeList = xmlDocResults.SelectNodes("//result[0]/address_component")

            For Each xmlAddressComponentNode In xmlAddressComponentNodeList
                Select Case xmlAddressComponentNode.SelectSingleNode("type[0]").Text
                    Case "street_number":
                        ret.street_number = xmlAddressComponentNode.SelectSingleNode("long_name").Text
                    Case "route":
                        ret.route_long = xmlAddressComponentNode.SelectSingleNode("long_name").Text
                        ret.route_short = xmlAddressComponentNode.SelectSingleNode("short_name").Text
                    Case "neighborhood":
                        ret.neighborhood = xmlAddressComponentNode.SelectSingleNode("long_name").Text
                    Case "locality":
                        ret.city = xmlAddressComponentNode.SelectSingleNode("long_name").Text
                    Case "administrative_area_level_2":
                        ret.county = xmlAddressComponentNode.SelectSingleNode("long_name").Text
                    Case "administrative_area_level_1":
                        ret.state_long = xmlAddressComponentNode.SelectSingleNode("long_name").Text
                        ret.state_short = xmlAddressComponentNode.SelectSingleNode("short_name").Text
                    Case "postal_code":
                        ret.postal_code = xmlAddressComponentNode.SelectSingleNode("long_name").Text
                    Case "postal_code_suffix":
                        ret.postal_code_suffix = xmlAddressComponentNode.SelectSingleNode("long_name").Text
                End Select
            Next xmlAddressComponentNode

            Set xmlAddressComponentNode = Nothing
            Set xmlAddressComponentNodeList = Nothing
            
            With xmlDocResults.SelectSingleNode("//result[0]/geometry/location")
                ret.geometry_lat = .SelectSingleNode("lat").Text
                ret.geometry_lon = .SelectSingleNode("lng").Text
            End With
            
            ret.location_type = xmlDocResults.SelectSingleNode("//result[0]/geometry/location_type").Text

        Case "ZERO_RESULTS"   'The geocode was successful but returned no results.
            ret.status_description = "The address probably not exists"

        Case "OVER_QUERY_LIMIT" 'The requestor has exceeded the limit of 2500 request/day.
            ret.status_description = "Requestor has exceeded the server limit"

        Case "REQUEST_DENIED"   'The API did not complete the request.
            ret.status_description = "Server denied the request"

        Case "INVALID_REQUEST"  'The API request is empty or is malformed.
            ret.status_description = "Request was empty or malformed"

        Case "UNKNOWN_ERROR"    'Indicates that the request could not be processed due to a server error.
            ret.status_description = "Unknown error"

        Case Else   'Just in case...
            ret.status_description = "Error"

    End Select

    
    Google_GeoCode = ret
    
    
    'In case of error, release the objects.
errorHandler:
    Set xmlStatusNode = Nothing
    Set xmlHttpReq = Nothing

End Function
Public Function Google_GeoCode_Reverse(lat As Double, lon As Double) As Google_GeoCode_Result


    Dim xmlHttpReq As Object 'New XMLHTTP30
    Dim xmlDocResults As Object 'DOMDocument30
    Dim xmlStatusNode As Object 'IXMLDOMNode


    Set xmlHttpReq = CreateObject("Microsoft.XMLHTTP")

    'On Error GoTo errorHandler


    '&key=AIzaSyDWrKMGCwpGG299G6rLMF5FnMuIZMugpoE

    'Send the request to the Google server.
    xmlHttpReq.Open "GET", "https://maps.googleapis.com/maps/api/geocode/xml?latlng=" & lat & "," & lon, False
    xmlHttpReq.Send

    'Read the results from the request.
    Set xmlDocResults = xmlHttpReq.responseXML
    'Results.LoadXML Request.responseText

    'Get the status node value.
    Set xmlStatusNode = xmlDocResults.SelectSingleNode("//status")

    Dim ret As Google_GeoCode_Result
    
    
    ret.status_code = UCase(xmlStatusNode.Text)

    'Based on the status node result, proceed accordingly.
    Select Case ret.status_code
        Case "OK"   'The API request was successful. At least one geocode was returned.

            Dim xmlAddressComponentNode As Object 'IXMLDOMNode
            Dim xmlAddressComponentNodeList As Object 'IXMLDOMNodeList


            Set xmlAddressComponentNodeList = xmlDocResults.SelectNodes("//result[0]/address_component")

            For Each xmlAddressComponentNode In xmlAddressComponentNodeList
                Select Case xmlAddressComponentNode.SelectSingleNode("type[0]").Text
                    Case "street_number":
                        ret.street_number = xmlAddressComponentNode.SelectSingleNode("long_name").Text
                    Case "route":
                        ret.route_long = xmlAddressComponentNode.SelectSingleNode("long_name").Text
                        ret.route_short = xmlAddressComponentNode.SelectSingleNode("short_name").Text
                    Case "neighborhood":
                        ret.neighborhood = xmlAddressComponentNode.SelectSingleNode("long_name").Text
                    Case "sublocality_level_1":
                        ret.city = xmlAddressComponentNode.SelectSingleNode("long_name").Text
                    Case "administrative_area_level_2":
                        ret.county = xmlAddressComponentNode.SelectSingleNode("long_name").Text
                    Case "administrative_area_level_1":
                        ret.state_long = xmlAddressComponentNode.SelectSingleNode("long_name").Text
                        ret.state_short = xmlAddressComponentNode.SelectSingleNode("short_name").Text
                    Case "postal_code":
                        ret.postal_code = xmlAddressComponentNode.SelectSingleNode("long_name").Text
                    Case "postal_code_suffix":
                        ret.postal_code_suffix = xmlAddressComponentNode.SelectSingleNode("long_name").Text
                End Select
            Next xmlAddressComponentNode

            Set xmlAddressComponentNode = Nothing
            Set xmlAddressComponentNodeList = Nothing


        Case "ZERO_RESULTS"   'The geocode was successful but returned no results.
            ret.status_description = "The address probably not exists"

        Case "OVER_QUERY_LIMIT" 'The requestor has exceeded the limit of 2500 request/day.
            ret.status_description = "Requestor has exceeded the server limit"

        Case "REQUEST_DENIED"   'The API did not complete the request.
            ret.status_description = "Server denied the request"

        Case "INVALID_REQUEST"  'The API request is empty or is malformed.
            ret.status_description = "Request was empty or malformed"

        Case "UNKNOWN_ERROR"    'Indicates that the request could not be processed due to a server error.
            ret.status_description = "Unknown error"

        Case Else   'Just in case...
            ret.status_description = "Error"

    End Select

    
    Google_GeoCode_Reverse = ret
    
    
    'In case of error, release the objects.
errorHandler:
    Set xmlStatusNode = Nothing
    Set xmlHttpReq = Nothing

End Function


Public Function Google_Elevation_SampledPath_1D(lat1 As Double, lon1 As Double, lat2 As Double, lon2 As Double, Optional samples As Integer = 20) As Variant

    'Dim lat1 As Double, lon1 As Double, lat2 As Double, lon2 As Double
    'Dim samples As Integer
    
    'Dim c1 As Variant, c2 As Variant
    
    'samples = 20
    
    'c1 = Array(40.020111, -75.217958)
    'c2 = Array(40.0172, -75.2408)
    
    'lat1 = c1(0): lon1 = c1(1)
    'lat2 = c2(0): lon2 = c2(1)


    Dim xmlHttpReq As Object 'New XMLHTTP30
    Dim xmlDocResults As Object  'DOMDocument30
    Dim xmlStatusNode As Object  'IXMLDOMNode


    Set xmlHttpReq = CreateObject("Microsoft.XMLHTTP")

    'On Error GoTo errorHandler


    'Send the request to the Google server.
    Dim requestURI As String
    
    requestURI = "https://maps.googleapis.com/maps/api/elevation/xml?samples=" & samples & "&path=" & lat1 & "," & lon1 & "|" & lat2 & "," & lon2
    
    Debug.Print requestURI
    
    
    xmlHttpReq.Open "GET", requestURI, False
    xmlHttpReq.Send

    'Read the results from the request.
    Set xmlDocResults = xmlHttpReq.responseXML
    'Results.LoadXML Request.responseText

    'Get the status node value.
    Set xmlStatusNode = xmlDocResults.SelectSingleNode("//status")

    
    Dim ret As Variant
    Dim elevationProfile() As Double
    Dim elevationProfileIdx As Long
    
    
    'ret.status_code = UCase(xmlStatusNode.text)

    'Based on the status node result, proceed accordingly.
    Select Case UCase(xmlStatusNode.Text)
        Case "OK"   'The API request was successful.

            Dim xmlResultNode As Object 'IXMLDOMNode
            Dim xmlResultNodeList As Object 'IXMLDOMNodeList

            Set xmlResultNodeList = xmlDocResults.SelectNodes("//result/elevation")
            
            ReDim elevationProfile(1 To xmlResultNodeList.length) As Double
            elevationProfileIdx = 1
            
            For Each xmlResultNode In xmlResultNodeList
               elevationProfile(elevationProfileIdx) = CDbl(xmlResultNode.Text)
               elevationProfileIdx = elevationProfileIdx + 1
            Next xmlResultNode
            

            Set xmlResultNode = Nothing
            Set xmlResultNodeList = Nothing
            
            ret = elevationProfile


        Case "ZERO_RESULTS"   'The geocode was successful but returned no results.
            ret = "The address probably not exists"

        Case "OVER_QUERY_LIMIT" 'The requestor has exceeded the limit of 2500 request/day.
            ret = "Requestor has exceeded the server limit"

        Case "REQUEST_DENIED"   'The API did not complete the request.
            ret = "Server denied the request"

        Case "INVALID_REQUEST"  'The API request is empty or is malformed.
            ret = "Request was empty or malformed"

        Case "UNKNOWN_ERROR"    'Indicates that the request could not be processed due to a server error.
            ret = "Unknown error"

        Case Else   'Just in case...
            ret = "Error"

    End Select

    
    Google_Elevation_SampledPath_1D = ret
    
    
    'In case of error, release the objects.
errorHandler:
    Set xmlStatusNode = Nothing
    Set xmlHttpReq = Nothing

End Function



