VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "gisPolygon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Public Points As Collection
Public Edges As Collection
Public BoundingBox As gisGeoBounds
Public Projection as gisProjection

'Private events
Private Sub Class_Initialize()
    Set BoundingBox = New clsGeoBounds
    Set obj.Points = New Collection
    Set obj.Edges = New Collection
End Sub
Private Sub Class_Terminate()
End Sub




'Public Methods
Public Function Create() As gisPolygon
    '...
End Function
Public Function CreateFrom1DArray(ByVal v as variant) as gisPolygon
    Dim obj as gisPolygon
    set obj = new gisPolygon
    With obj.points
        For i = LBound(v) To UBound(v)
            .Add clsGeoPoint.Create(v(i))
        Next i
    End with
End Function
Public Function CreateFromCollection(ByVal col as collection) as gisPolygon
    Dim obj as gisPolygon
    set obj = new gisPolygon
    If TypeName(Params(0)) = "Collection" Then
        Dim v as variant
        With obj.Points
            For each v in col
                .Add clsGeoPoint.Create(col)
            Next i
        End with 
    End If
End Function
Public Function CreateFromPointString(ByVal sPnts as string) as gisPolygon   'Assume "[[x1,y1],[x2,y2],[x3,y3],...]"
    Dim obj as gisPolygon
    set obj = new gisPolygon
    Dim data As Variant
    data = Replace(sPnts, "[", "")
    data = Split(Mid(data, 1, Len(data) - 2), "],")
    
    Dim sData As Variant
    For i = LBound(data) To UBound(data)
        sData = Split(data(i), ",")
        obj.Points.Add clsGeoPoint.CreateFromXY(CDbl(sData(0)), CDbl(sData(1)))
    Next
End Function







Public Function Contains(ByVal obj As clsGeoPoint) As Boolean
    'Optimise contains algorithm based on bounding box
    If Not Me.BoundingBox.Contains(obj) Then
        Countains = False
        Exit Function
    End If
    
    'See http://geomalgorithms.com/a03-_inclusion.html
    'Uses 'Crossing' algorithm
    Dim edge As clsGeoLine
    Dim crossingLine As clsGeoLine
    Set crossingLine = clsGeoLine.From(0, 0, obj.x, obj.y)
    Dim crossingNo As Long
    
    'Count number of intersects
    For Each edge In Me.Edges
        If edge.Intersects(crossingLine) Then crossingNo = crossingNo + 1
    Next
    
    'If number of intersects is odd, then point is within, else it's outside
    Contains = crossingNo And 1
End Function

Public Sub zTEST__Contains()
    Dim testPolygons As Collection
    Dim subCol As Collection
    Set testPolygons = New Collection
   
    testPolygons.Add clsGeoPolygon.From("[[989960.23,991022.49],[991096.86,991022.49],[991096.86,989999.55],[989962.87,989999.55],[989960.23,991022.49],[990374.14,990721.96],[990712.49,990721.96],[990712.49,990352.86],[990374.93,990352.86],[990374.14,990721.96]]"), "Hole"
    testPolygons.Add clsGeoPolygon.From("[[990344.99,990999.62],[989688.01,990999.62],[989688.05,989970.79],[990344.99,989970.35],[990344.99,990999.62],[991936.33,991010.54],[990731.87,991006.92],[990735.28,989999.55],[991936.33,989999.55],[991936.33,991010.54]]"), "Combined"
    testPolygons.Add clsGeoPolygon.From("[[990013.51,991000.18],[991001.91,991000.18],[991001.91,990000.18],[990001.88,990000.18]]"), "Complete"
   
    Dim testPoints As New Collection
    testPoints.Add clsGeoPoint.From("[990558.57,990561.81]"), "Complete only"
    testPoints.Add clsGeoPoint.From("[990171.68,990561.81]"), "All"
    testPoints.Add clsGeoPoint.From("[990558.57,990182.22]"), "Complete&Hole"
    
    Debug.Print iff(testPolygons("Complete").Contains(testPoints("Complete only")), "PASS", "FAIL")
    Debug.Print iff(Not testPolygons("Combined").Contains(testPoints("Complete only")), "PASS", "FAIL")
    Debug.Print iff(Not testPolygons("Hole").Contains(testPoints("Complete only")), "PASS", "FAIL")
    
    Debug.Print iff(testPolygons("Complete").Contains(testPoints("All")), "PASS", "FAIL")
    Debug.Print iff(testPolygons("Combined").Contains(testPoints("All")), "PASS", "FAIL")
    Debug.Print iff(testPolygons("Hole").Contains(testPoints("All")), "PASS", "FAIL")
    
    Debug.Print iff(testPolygons("Complete").Contains(testPoints("Complete&Hole")), "PASS", "FAIL")
    Debug.Print iff(Not testPolygons("Combined").Contains(testPoints("Complete&Hole")), "PASS", "FAIL")
    Debug.Print iff(testPolygons("Hole").Contains(testPoints("Complete&Hole")), "PASS", "FAIL")
End Sub



'Draws the polygon to a specified sheet range
Public Sub Draw(sheet As Worksheet, Optional rSize As Range)
    
End Sub
Private Function iff(ByVal cond As Boolean, ByVal a As Variant, ByVal b As Variant) As Variant
    iff = b
    If cond Then iff = a
End Function









'NOTES:
'
'TO EXPORT POLYGONS FROM MAPINFO USE THIS MAPBASIC:
'(Exports table1 with COL1 followed by the point array)
'------------------------------------------------------------------------
'Dim aAliasObj, aAliasCol As alias
'Dim s, sKey As String
'Dim oObject As Object
'Dim iIndex, iPnt, iPoly As Integer
'set coordsys table TableInfo(1,1)
'For iIndex = 1 To TableInfo(1, 8)
'  fetch rec iIndex from TableInfo(1,1)
'  aAliasObj = TableInfo(1, 1) & ".obj"
'  aAliasCol = TableInfo(1, 1) & ".Col1"
'  oObject = aAliasObj
'  sKey = aAliasCol
'
'  s = s & sKey & "|" & "["
'  For iPnt = 1 To ObjectInfo(oObject, 22) - 1 'loop through all points of polygon 1.
'    If iPnt = 1 Then
'      s = s & "[" & ObjectNodeX(oObject, 1, iPnt) & "," & ObjectNodeY(oObject, 1, iPnt) & "]"
'    Else
'      s = s & ",[" & ObjectNodeX(oObject, 1, iPnt) & "," & ObjectNodeY(oObject, 1, iPnt) & "]"
'    End If
'  Next
'  For iPoly = 2 To ObjectInfo(oObject, 21)
'    s = s & ",[" & ObjectNodeX(oObject, 1, 1) & "," & ObjectNodeY(oObject, 1, 1) & "]"
'    For iPnt = 1 To ObjectInfo(oObject, 20 + iPoly) 'loop through all points of polygon 1.
'      s = s & ",[" & ObjectNodeX(oObject, iPoly, iPnt) & "," & ObjectNodeY(oObject, iPoly, iPnt) & "]"
'    Next
'  Next
'  s = s & "]" & Chr$(10)
'Next
'Print s

