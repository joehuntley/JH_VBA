Attribute VB_Name = "JH_VZW_OPS_Helper"
Option Explicit

Private Type SiteLevel_Site
    init_flag As Boolean
    SiteName As String
    Fields As Dictionary
End Type


Private Type SiteLevel_Tech
    init_flag As Boolean
    TechName As String
    Sites() As SiteLevel_Site
End Type

Private Type SiteLevel_Zone
    init_flag As Boolean
    CalloutZone As String
    Techs() As SiteLevel_Tech
End Type

Private Type SiteLevel_Mgr
    init_flag As Boolean
    ManagerName As String
    Zones() As SiteLevel_Zone
End Type


Type KML_Helper_CreateTieredKML_Options
    StyleRotationColors() As Variant
    StyleRotationColor_Column As Integer
    StyleRotationColor_Index As Integer
    
    DefaultLatitude As Double
    DefaultLongitude As Double
End Type

Public Sub VZW_OPS_CreateKML_SitesByMgrTechZone()




    Dim outFile As String: outFile = ThisWorkbook.Path & "\sites_by_mgr.kml"
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("Sites")
    Dim pivotColumns() As Variant: pivotColumns = Array("MGR_NAME", "CALLOUT_ZONE", "TECH_NAME")
    Dim headerRow As Long: headerRow = 1
    Dim kml_options As KML_Helper_CreateTieredKML_Options
    
    
    
    Dim C_SITE As Long: C_SITE = Worksheet_GetColumnByName(ws, "SITE_NAME")
    Dim C_MGR As Long: C_MGR = Worksheet_GetColumnByName(ws, "MGR_NAME")
    Dim C_TECH As Long: C_TECH = Worksheet_GetColumnByName(ws, "TECH_NAME")
    Dim C_ZONE As Long: C_ZONE = Worksheet_GetColumnByName(ws, "CALLOUT_ZONE")
    
    Dim ssMap As SpreadsheetTableToPivotMapping

    ssMap = SpreadsheetTableToPivotMap(ws, pivotColumns)
    
    Dim i As Long, j As Long, k As Long, l As Long
    
    Dim mgrCount As Long: mgrCount = Array_Count(ssMap.pivotToRowMap)
    Dim zoneCount As Long, techCount As Long, siteCount As Long
    
    Dim SitesByMgr() As SiteLevel_Mgr
    ReDim SitesByMgr(1 To mgrCount) As SiteLevel_Mgr
    
    ' Top Level: Mgr
    For i = 1 To mgrCount
        zoneCount = Array_Count(ssMap.pivotToRowMap(i))
        ReDim SitesByMgr(i).Zones(1 To zoneCount) As SiteLevel_Zone
        
        For j = 1 To zoneCount
            techCount = Array_Count(ssMap.pivotToRowMap(i)(j))
            ReDim SitesByMgr(i).Zones(j).Techs(1 To techCount) As SiteLevel_Tech
            
            For k = 1 To techCount
                siteCount = Array_Count(ssMap.pivotToRowMap(i)(j)(k))
                ReDim SitesByMgr(i).Zones(j).Techs(k).Sites(1 To siteCount) As SiteLevel_Site
            Next k
        Next j
    Next i
    
    Dim row As Long, lastRow As Long
    
    Dim colDict As Dictionary
    Dim fieldName As Variant, fieldVal As Variant
    
    Set colDict = Worksheet_GetColumnDictionary(ws, headerRow)
    lastRow = Worksheet_GetLastRow(ws)
    

    
    For row = 2 To lastRow
        i = ssMap.RowToPivotMap(row, 1) ' Top Level Index
        j = ssMap.RowToPivotMap(row, 2) ' Second Level Index
        k = ssMap.RowToPivotMap(row, 3) ' Third Level Index
        l = ssMap.RowToPivotMap(row, 4) ' Item

        ' Fill in top level values
        If SitesByMgr(i).init_flag = False Then
            SitesByMgr(i).init_flag = True
            SitesByMgr(i).ManagerName = ws.Cells(row, C_MGR)
        End If
        ' Fill in second level values
        If SitesByMgr(i).Zones(j).init_flag = False Then
            SitesByMgr(i).Zones(j).init_flag = True
            SitesByMgr(i).Zones(j).CalloutZone = ws.Cells(row, C_ZONE)
        End If
        ' Fill in third level values
        If SitesByMgr(i).Zones(j).Techs(k).init_flag = False Then
            SitesByMgr(i).Zones(j).Techs(k).init_flag = True
            SitesByMgr(i).Zones(j).Techs(k).TechName = ws.Cells(row, C_TECH)
        End If
        
        ' Fill in item values
        SitesByMgr(i).Zones(j).Techs(k).Sites(l).SiteName = ws.Cells(row, C_SITE)
        
        Set SitesByMgr(i).Zones(j).Techs(k).Sites(l).Fields = New Dictionary
        
        For Each fieldName In colDict
            SitesByMgr(i).Zones(j).Techs(k).Sites(l).Fields.Add Key:=fieldName, Item:=ws.Cells(row, colDict(fieldName))
        Next fieldName
    Next row


    Dim kml As String
    
    
    
    kml = "<?xml version=""1.0"" encoding=""UTF-8""?>" & vbCrLf & _
          "<kml xmlns=""http://www.opengis.net/kml/2.2"">" & vbCrLf & _
          "<Document>" & vbCrLf & "<name>Test</name>" & vbCrLf
          
          
    kml = kml & "<Style id=""baseStyle"">" & vbCrLf & _
                "    <IconStyle>" & vbCrLf & _
                "        <color>ffffffff</color><scale>1.0</scale>" & vbCrLf & _
                "        <Icon><href>http://maps.google.com/mapfiles/kml/paddle/wht-blank.png</href></Icon>" & vbCrLf & _
                "        <hotSpot x=""32"" y=""1"" xunits=""pixels"" yunits=""pixels""/>" & vbCrLf & _
                "    </IconStyle>" & vbCrLf & _
                "    <LabelStyle><scale>0.5</scale></LabelStyle>" & vbCrLf & _
                "</Style>"
    
    
    Dim indent As Long
    Dim styleColors() As Variant
    Dim styleIdx As Long
    
    
    styleColors = Array( _
        "ffff0000", "ff00ff00", "ff0000ff", "ffff00ff", _
        "ffff6600", "ffffcccc", "ff996633", "ffffcc00", "ffffff99", _
        "ffff6600", "ff99ff99", "ff00cccc", "ff00ccff", "ff9900cc", _
        "ff9999ff", "ffff0099", "ffcc0099", "ffffffff" _
    )
    styleIdx = 0
    
    indent = 1

    For i = LBound(SitesByMgr) To UBound(SitesByMgr)
        kml = kml & Space(indent * 4) & "<Folder>" & vbCrLf
        
        indent = indent + 1
        kml = kml & Space(indent * 4) & "<open>0</open>" & vbCrLf & _
                    Space(indent * 4) & "<name>" & SitesByMgr(i).ManagerName & "</name>" & vbCrLf
        
        For j = LBound(SitesByMgr(i).Zones) To UBound(SitesByMgr(i).Zones)
            kml = kml & Space(indent * 4) & "<Folder>" & vbCrLf
            
            indent = indent + 1
            kml = kml & Space(indent * 4) & "<open>0</open>" & vbCrLf & _
                        Space(indent * 4) & "<name>" & SitesByMgr(i).Zones(j).CalloutZone & "</name>" & vbCrLf
                        
                        
            For k = LBound(SitesByMgr(i).Zones(j).Techs) To UBound(SitesByMgr(i).Zones(j).Techs)
                kml = kml & Space(indent * 4) & "<Folder>" & vbCrLf
                
                indent = indent + 1
                kml = kml & Space(indent * 4) & "<open>0</open>" & vbCrLf & _
                            Space(indent * 4) & "<name>" & SitesByMgr(i).Zones(j).Techs(k).TechName & "</name>" & vbCrLf
    
            
                For l = LBound(SitesByMgr(i).Zones(j).Techs(k).Sites) To UBound(SitesByMgr(i).Zones(j).Techs(k).Sites)
                    kml = kml & Space(indent * 4) & "<Placemark>" & vbCrLf
                    
                    indent = indent + 1
                    kml = kml & Space(indent * 4) & "<name>" & SitesByMgr(i).Zones(j).Techs(k).Sites(l).SiteName & "</name>" & vbCrLf & _
                                Space(indent * 4) & "<visibility>1</visibility>" & vbCrLf & _
                                Space(indent * 4) & "<styleUrl>#baseStyle</styleUrl>" & vbCrLf
                        
                    kml = kml & Space(indent * 4) & "<Style id=""inline""><IconStyle><color>" & styleColors(styleIdx) & "</color><colorMode>normal</colorMode></IconStyle></Style>" & vbCrLf


                    
                    With SitesByMgr(i).Zones(j).Techs(k).Sites(l)
                        Dim lat As Double, lon As Double
                        lat = .Fields("LATITUDE"): lon = .Fields("LONGITUDE")
                        
                        If Abs(lat) < 1 Or Abs(lon) < 0 Then
                            lat = 28.371219: lon = -91.437128
                        End If
                    
                        kml = kml & Space(indent * 4) & "<Point><coordinates>" & lon & "," & lat & "</coordinates></Point>" & vbCrLf
                    
                        kml = kml & Space(indent * 4) & "<ExtendedData>" & vbCrLf
                        indent = indent + 1
                        
                        For Each fieldName In .Fields
                            kml = kml & Space(indent * 4) & "<Data name=""" & fieldName & """><value>" & .Fields(fieldName) & "</value></Data>" & vbCrLf
                        Next fieldName
                    End With
                    indent = indent - 1
                    kml = kml & Space(indent * 4) & "</ExtendedData>" & vbCrLf
                    
                    indent = indent - 1
                    kml = kml & Space(indent * 4) & "</Placemark>" & vbCrLf
                Next l
                
                styleIdx = styleIdx + 1
                styleIdx = styleIdx Mod Array_Count(styleColors)
                
                indent = indent - 1
                kml = kml & Space(indent * 4) & "</Folder>" & vbCrLf
            Next k
            
            indent = indent - 1
            kml = kml & Space(indent * 4) & "</Folder>" & vbCrLf
        Next j
        
        indent = indent - 1
        kml = kml & Space(indent * 4) & "</Folder>" & vbCrLf
    Next i
    
    kml = kml & "</Document></kml>"

    
    Set colDict = Nothing
    
    Dim fH As Long: fH = FreeFile
    
    Open outFile For Output Access Write As fH
        Print #fH, kml
    Close fH
    
End Sub



Public Function KML_Helper_CreateTieredKMLFile()



    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("Sites")
    Dim pivotColumns() As Variant: pivotColumns = Array("MGR_NAME", "CALLOUT_ZONE", "TECH_NAME")
    Dim headerRow As Long: headerRow = 1
    Dim outFile As String: outFile = ThisWorkbook.Path & "\test2.kml"
    Dim includeKmlHeaderFooter As Boolean
    
    Dim kml_options As KML_Helper_CreateTieredKML_Options
    
    includeKmlHeaderFooter = Len(outFile) > 0
    
    Dim groupByColumns() As String
    groupByColumns = Array_ToString(Array("MGR_NAME", "CALLOUT_ZONE", "TECH_NAME", "SITE_NAME"))
    
    Debug.Print
    
    
    Dim kml As String
    
    Dim C_SITE As Long: C_SITE = Worksheet_GetColumnByName(ws, "SITE_NAME")
    Dim C_MGR As Long: C_MGR = Worksheet_GetColumnByName(ws, "MGR_NAME")
    Dim C_TECH As Long: C_TECH = Worksheet_GetColumnByName(ws, "TECH_NAME")
    Dim C_ZONE As Long: C_ZONE = Worksheet_GetColumnByName(ws, "CALLOUT_ZONE")
    
    Dim ssMap As SpreadsheetTableToPivotMapping

    ssMap = SpreadsheetTableToPivotMap(ws, pivotColumns)
    
    'Dim i As Long, j As Long, k As Long, l As Long
    
    
    'Dim row As Long, lastRow As Long
    
    'Dim colDict As Dictionary
    'Dim fieldName As Variant, fieldVal As Variant
   
    If includeKmlHeaderFooter Then
        kml = "<?xml version=""1.0"" encoding=""UTF-8""?>" & vbCrLf & _
              "<kml xmlns=""http://www.opengis.net/kml/2.2"">" & vbCrLf & _
              "<Document>" & vbCrLf & "<name>Test</name>" & vbCrLf
    End If
          
    kml = kml & "<Style id=""baseStyle"">" & vbCrLf & _
                "    <IconStyle>" & vbCrLf & _
                "        <color>ffffffff</color><scale>1.0</scale>" & vbCrLf & _
                "        <Icon><href>http://maps.google.com/mapfiles/kml/paddle/wht-blank.png</href></Icon>" & vbCrLf & _
                "        <hotSpot x=""32"" y=""1"" xunits=""pixels"" yunits=""pixels""/>" & vbCrLf & _
                "    </IconStyle>" & vbCrLf & _
                "    <LabelStyle><scale>0.5</scale></LabelStyle>" & vbCrLf & _
                "</Style>" & vbCrLf
    
    
    Dim indent As Integer
    Dim rotateStyleColors() As Variant
    Dim rotateStyleColorColumn As Long
    
    
    rotateStyleColors = Array("ffff0000", "ff00ff00", "ff0000ff", "ffff00ff", _
        "ffff6600", "ffffcccc", "ff996633", "ffffcc00", "ffffff99", _
        "ffff6600", "ff99ff99", "ff00cccc", "ff00ccff", "ff9900cc", _
        "ff9999ff", "ffff0099", "ffcc0099", "ffffffff" _
    )
    rotateStyleColorColumn = 2
    
    
    kml = kml & KML_Helper_CreateTieredKMLFile_BuildFolderKML( _
        ws:=ws, _
        groupByColumnNames:=groupByColumns, _
        pivotToRowMapArr:=ssMap.pivotToRowMap, _
        currentGroupByCol:=LBound(groupByColumns), _
        headerRow:=headerRow, indent:=indent _
    )
    
    
    
    If includeKmlHeaderFooter Then kml = kml & "</Document></kml>"

    
    'Set colDict = Nothing
    
    If Len(outFile) > 0 Then
        Dim fH As Long: fH = FreeFile
        
        Open outFile For Output Access Write As fH
            Print #fH, kml
        Close fH
    Else
        KML_Helper_CreateTieredKMLFile = kml
    End If
    
End Function

Public Function KML_Helper_CreateTieredKMLFile_BuildFolderKML(ws As Worksheet, groupByColumnNames() As String, pivotToRowMapArr() As Variant, Optional currentGroupByCol As Integer = 1, Optional headerRow As Long = 1, Optional indent As Integer = 0) As String

    Debug.Assert currentGroupByCol >= LBound(groupByColumnNames)
    Debug.Assert currentGroupByCol <= UBound(groupByColumnNames)

    
    Dim i As Long
    Dim kml As String
    Dim firstRow As Long, row As Long
    Dim folderName As String
    Dim colDict As Dictionary
    
    Dim arrTmp() As Variant
    
    
    Set colDict = Worksheet_GetColumnDictionary(ws, headerRow)
    'colDict.CompareMode = vbTextCompare
    
    kml = kml & Space(indent * 4) & "<Folder>" & vbCrLf
    
    
    firstRow = KML_Helper_CreateTieredKMLFile_FindFirstGroupRow(groupByColumnNames, pivotToRowMapArr, currentGroupByCol, headerRow)
    
    folderName = ws.Cells(firstRow, colDict(groupByColumnNames(currentGroupByCol)))
    
    
    indent = indent + 1
    kml = kml & Space(indent * 4) & "<name>" & folderName & "</name>" & vbCrLf & _
                Space(indent * 4) & "<open>0</open>" & vbCrLf
                
    If currentGroupByCol < UBound(groupByColumnNames) Then
        ' Not at last level: build folder structure recursively.
        
        For i = 1 To UBound(pivotToRowMapArr) ' We use 1 for L-Bound because 0 may or may not exist but if it does, only contains the count (see: SpreadsheetTableToPivotMap)
            Debug.Assert IsArray(pivotToRowMapArr(i)) = True
            
            arrTmp = pivotToRowMapArr(i) ' Need to use a temp variable here
            kml = kml & KML_Helper_CreateTieredKMLFile_BuildFolderKML( _
                ws:=ws, _
                groupByColumnNames:=groupByColumnNames, _
                pivotToRowMapArr:=arrTmp, _
                currentGroupByCol:=currentGroupByCol + 1, _
                headerRow:=headerRow, indent:=indent _
            )
        Next i
    Else
       ' Last level: build KML for placemarks
       
        For i = 1 To UBound(pivotToRowMapArr) ' We use 1 for L-Bound because 0 may or may not exist but if it does, only contains the count (see: SpreadsheetTableToPivotMap)
            Debug.Assert IsArray(pivotToRowMapArr(i)) = False
            
            row = pivotToRowMapArr(i)
            
            
            Dim lat As Double, lon As Double
            lat = ws.Cells(row, colDict("LATITUDE"))
            lon = ws.Cells(row, colDict("LONGITUDE"))
            
            If lat = 0 And lon = 0 Then
                lat = 28.371219: lon = -91.437128
            End If
            
            kml = kml & Space(indent * 4) & "<Placemark>" & vbCrLf
            
            indent = indent + 1
            kml = kml & Space(indent * 4) & "<name>" & ws.Cells(row, colDict(groupByColumnNames(currentGroupByCol))) & "</name>" & vbCrLf & _
                        Space(indent * 4) & "<visibility>1</visibility>" & vbCrLf & _
                        Space(indent * 4) & "<styleUrl>#baseStyle</styleUrl>" & vbCrLf
            
            'kml = kml & Space(indent * 4) & "<Style id=""inline""><IconStyle><color>" & styleColors(styleIdx) & "</color><colorMode>normal</colorMode></IconStyle></Style>" & vbCrLf
            
            
        
            kml = kml & Space(indent * 4) & "<Point><coordinates>" & lat & "," & lon & "</coordinates></Point>" & vbCrLf
        
            
            kml = kml & Space(indent * 4) & "<ExtendedData>" & vbCrLf
            indent = indent + 1
            
            Dim fieldName As Variant
            For Each fieldName In colDict
                kml = kml & Space(indent * 4) & "<Data name=""" & fieldName & """><value>" & ws.Cells(row, colDict(fieldName)) & "</value></Data>" & vbCrLf
            Next fieldName
            indent = indent - 1
            kml = kml & Space(indent * 4) & "</ExtendedData>" & vbCrLf
            
            indent = indent - 1
            kml = kml & Space(indent * 4) & "</Placemark>" & vbCrLf
        Next i
    End If
    
    indent = indent - 1
    kml = kml & Space(indent * 4) & "</Folder>" & vbCrLf
    
    Set colDict = Nothing
    
    
    KML_Helper_CreateTieredKMLFile_BuildFolderKML = kml

End Function
Public Function KML_Helper_CreateTieredKMLFile_FindFirstGroupRow(groupByColumnNames() As String, pivotToRowMapArr() As Variant, currentGroupByCol As Integer, Optional headerRow As Long = 1) As Long

    Debug.Assert currentGroupByCol >= LBound(groupByColumnNames)
    Debug.Assert currentGroupByCol <= UBound(groupByColumnNames)
    
    Dim arrTmp() As Variant
    Dim ret As Long
                
    If currentGroupByCol < UBound(groupByColumnNames) Then
        ' Not at last level
        Debug.Assert IsArray(pivotToRowMapArr(1)) = True
            
        arrTmp = pivotToRowMapArr(1) ' Need to use a temp variable here
        ret = KML_Helper_CreateTieredKMLFile_FindFirstGroupRow(groupByColumnNames, arrTmp, currentGroupByCol + 1, headerRow)
    Else
        ' Last level
        Debug.Assert IsArray(pivotToRowMapArr(1)) = False
        ret = pivotToRowMapArr(1)
    End If
    
    KML_Helper_CreateTieredKMLFile_FindFirstGroupRow = ret

End Function
