VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "gisFeatureCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type BBox
  MinLat As Double
  MaxLat As Double
  MinLng As Double
  MaxLng As Double
End Type

Private sData As String
Private pBounds As BBox


Public Function Create() As gisFeatureCollection
  Set Create = New gisFeatureCollection
  Call Create.protInit
End Function

'Add a point to the layer
'@param {gisIPoint} Point to add
'@param {Dictionary} Properties of the point. Note properties can be added to give markers a custom style. marker-color:"#ff0000", marker-size: "small"/"large"/"medium" and marker-symbol: "..."
Public Sub AddPoint(ByVal point As gisIPoint, Optional ByVal Data As Object)
  Call AddPoint2(point.Latitude, point.Longitude, Data)
End Sub

'Add a point to the layer
'@param {Double} Latitude
'@param {Double} Longitude
'@param {Dictionary} Properties of the point. Note properties can be added to give markers a custom style. marker-color:"#ff0000", marker-size: "small"/"large"/"medium" and marker-symbol: "..."
Public Sub AddPoint2(ByVal Latitude As Double, ByVal Longitude As Double, Optional ByVal Data As Object)
  Call updateBounds(Array(Latitude, Longitude))
  Dim sPoint As String: sPoint = "{""type"": ""Feature"",""properties"": " & serialiseDict(Data) & ",""geometry"": {""type"": ""Point"",""coordinates"": [" & Longitude & "," & Latitude & "]}}"
  sData = sData & IIf(sData = "", "", "," & vbCrLf) & sPoint
End Sub


'Add a line to the layer
'@param {Double} Latitude  of start point
'@param {Double} Longitude of start point
'@param {Double} Latitude  of end point
'@param {Double} Longitude of end point
'@param {Dictionary} Properties of the point. Note properties can be added to give markers a custom style. stroke:"#ff0000", stroke-width: number and stroke-opacity: 0-1
Public Sub AddLine(ByVal oStart As gisIPoint, oEnd As gisIPoint, Optional ByVal Data As Object)
  Call AddLine2(oStart.Latitude, oStart.Longitude, oEnd.Latitude, oEnd.Longitude, Data)
End Sub

'Add a line to the layer
'@param {Double} Latitude  of start point
'@param {Double} Longitude of start point
'@param {Double} Latitude  of end point
'@param {Double} Longitude of end point
'@param {Dictionary} Properties of the point. Note properties can be added to give markers a custom style. stroke:"#ff0000", stroke-width: number and stroke-opacity: 0-1
Public Sub AddLine2(ByVal StartLat As Double, StartLong As Double, EndLat As Double, EndLong As Double, Optional ByVal Data As Object)
  Call updateBounds(Array(StartLat, StartLong, EndLat, EndLong))
  Dim sLine As String: sLine = "{""type"": ""Feature"",""properties"": " & serialiseDict(Data) & ",""geometry"": {""type"": ""LineString"",""coordinates"": [[" & StartLong & "," & StartLat & "],[" & EndLong & "," & EndLat & "]]}}"
  sData = sData & IIf(sData = "", "", "," & vbCrLf) & sLine
End Sub

'Add a line to the layer
'@param {Double} vPolyline  1D array of x,y pairs
'@param {Dictionary} Properties of the point. Note properties can be added to give markers a custom style. stroke:"#ff0000", stroke-width: number and stroke-opacity: 0-1
Public Sub AddPolygon(ByVal vPolygon As Variant, Optional ByVal Data As Object)
  Call updateBounds(vPolygon)
  Dim sArr As String: sArr = serialiseTo2DArray(vPolygon)
  Dim sLine As String: sLine = "{""type"": ""Feature"",""properties"": " & serialiseDict(Data) & ",""geometry"": {""type"": ""LineString"",""coordinates"": " & sArr & "}}"
  sData = sData & IIf(sData = "", "", "," & vbCrLf) & sLine
End Sub

'Add a line to the layer
'@param {Double} vPolyline  1D array of x,y pairs
'@param {Dictionary} Properties of the point. Note properties can be added to give markers a custom style. stroke:"#ff0000", stroke-width: number and stroke-opacity: 0-1
Public Sub AddPolyLine(ByVal vPolyline As Variant, Optional ByVal Data As Object)
  Call updateBounds(vPolyline)
  Dim sArr As String: sArr = serialiseTo2DArray(vPolyline)
  Dim sLine As String: sLine = "{""type"": ""Feature"",""properties"": " & serialiseDict(Data) & ",""geometry"": {""type"": ""Polygon"",""coordinates"": " & sArr & "}}"
  sData = sData & IIf(sData = "", "", "," & vbCrLf) & sLine
End Sub

'Returns Me
Public Property Get Object() As gisFeatureCollection
  Set Object = Me
End Property

'Returns GeoJSON string
Public Function ToString() As String
  ToString = "{""type"": ""FeatureCollection"", ""features"": [" & sData & "]}"
End Function

Public Property Get Bounds() As Variant
  With pBounds
    Bounds = Array(.MinLat, .MinLng, .MaxLat, .MaxLng)
  End With
End Property


Private Sub updateBounds(ByVal vPoints As Variant)
  Dim i As Long: For i = LBound(vPoints) To UBound(vPoints) Step 2
    If pBounds.MaxLat < vPoints(i) Then pBounds.MaxLat = vPoints(i)
    If pBounds.MinLat > vPoints(i) Then pBounds.MinLat = vPoints(i)
    If pBounds.MaxLng < vPoints(i + 1) Then pBounds.MaxLng = vPoints(i + 1)
    If pBounds.MinLng > vPoints(i + 1) Then pBounds.MinLng = vPoints(i + 1)
  Next
End Sub

'@param array of Longitude,Latitude pairs e.g. [lng1,lat1,lng2,lat2,lng3,lat3,lng4,lat4]
Private Function serialiseTo2DArray(ByVal v As Variant) As String
  Dim sData As String, i As Long
  For i = LBound(v) To UBound(v) Step 2
    sData = sData & IIf(sData = "", "", ",") & "[" & v(i) & "," & v(i + 1) & "]"
  Next
  serialiseTo2DArray = "[" & sData & "]"
End Function
Private Function serialiseDict(ByVal Data As Object) As String
  If Data Is Nothing Then
    serialiseDict = "{}"
    Exit Function
  End If
  
  Dim sData As String: sData = ""
  Dim vKey: For Each vKey In Data.keys()
    sData = sData & IIf(sData = "", "", ",") & serialise(vKey) & ":" & serialise(Data(vKey))
  Next
  serialiseDict = "{" & sData & "}"
End Function
Private Function serialise(ByVal v As Variant) As String
  If VarType(v) = vbString Then
    serialise = """" & v & """"
  Else
    serialise = v
  End If
End Function



Friend Sub protInit()
  pBounds.MaxLat = -1E+16
  pBounds.MaxLng = -1E+16
  pBounds.MinLat = 1E+16
  pBounds.MinLng = 1E+16
End Sub
