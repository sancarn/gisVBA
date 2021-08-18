VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "gisPoint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Implements gisIPoint

'Library: gisVBA_WIP
'URL:     https://github.com/sancarn/gisVBA/tree/dev/tbc
'Credits:
'* Initial version by Ordnance Survey ( https://www.ordnancesurvey.co.uk/business-government/tools-support/os-net/coordinates )
'* 2021-08-17 Adapted by Sancarn (This file has several versions, but many took inspiration from OS above. This file was re-written to fix some issues)
'             * Code has been expanded to cope with multiple coordinate systems.
'             * Class singleton with Create constructors for ease of use.
'             * Class singleton with Create constructors for ease of use.
'             * Tested against www.nearby.org.uk/tests/geotools2.js for validation


Private Enum ECoordType
  LatLng 'WGS84
  NGR_UK 'OSGB36
  NGR_Ireland '1965
  
  'Suggest keeping this at the end for dynamic arrays below
  UNK
End Enum

Private Const PI As Double = 3.14159265358979
Private pB(0 To ECoordType.UNK) As Boolean
Private pX(0 To ECoordType.UNK) As Double
Private pY(0 To ECoordType.UNK) As Double


Public Enum ENGRType
  BritishNationalGrid
  IrishNationalGrid
End Enum

'Create a point from Latitude and Longitude coordinates
'@param {Double}    fLatitude  - Latitude coord
'@param {Double}    fLongitude - Longitude coord
'@returns {gisPoint} A point object
Public Function Create(ByVal fLatitude As Double, ByVal fLongitude As Double) As gisPoint
  Set Create = New gisPoint
  Call Create.protInit(LatLng, fLongitude, fLatitude)
End Function

'Create a point from national grid coordinates
'@param {Double}    fEastings  - Eastings coord
'@param {Double}    fNorthings - Northings coord
'@param {ENGRType?} iNGRType   - BritishNationalGrid or IrishNationalGrid depending on coord typeAnd
'@returns {gisPoint} A point object
Public Function CreateFromProjection(ByVal fEastings As Double, ByVal fNorthings As Double, Optional ByVal iNGRType As ENGRType = BritishNationalGrid) As gisPoint
  Dim pnt As gisPoint: Set pnt = New gisPoint

  Dim lat1 As Double, lng1 As Double
  Dim X1 As Double, Y1 As Double, Z1 As Double
  Dim x2 As Double, Y2 As Double, Z2 As Double
  Dim fLatitude As Double, fLongitude As Double
  Select Case iNGRType
    Case BritishNationalGrid
      'Get base latitude and longitude
      height = 0
      lat1 = E_N_to_Lat(fEastings, fNorthings, 6377563.396, 6356256.91, 400000, -100000, 0.999601272, 49#, -2#)
      lng1 = E_N_to_Long(fEastings, fNorthings, 6377563.396, 6356256.91, 400000, -100000, 0.999601272, 49#, -2#)
      
      'Get cartesian coordinates
      X1 = Lat_Long_H_to_X(lat1, lng1, height, 6377563.396, 6356256.91)
      Y1 = Lat_Long_H_to_Y(lat1, lng1, height, 6377563.396, 6356256.91)
      Z1 = Lat_H_to_Z(lat1, height, 6377563.396, 6356256.91)
      
      'Offset coordinates to real coordinates
      x2 = Helmert_X(X1, Y1, Z1, 446.448, 0.247, 0.8421, -20.4894)
      Y2 = Helmert_Y(X1, Y1, Z1, -125.157, 0.1502, 0.8421, -20.4894)
      Z2 = Helmert_Z(X1, Y1, Z1, 542.06, 0.1502, 0.247, -20.4894)
      
      'Convert back to latitude, longitude
      fLatitude = XYZ_to_Lat(x2, Y2, Z2, 6378137#, 6356752.313)
      fLongitude = XYZ_to_Long(x2, Y2)

      'Initialise coord data
      Call pnt.protInit(LatLng, fLongitude, fLatitude)
      Call pnt.protInit(NGR_UK, fEastings, fNorthings)
    Case IrishNationalGrid
      'Get base latitude and longitude
      height = 0
      lat1 = E_N_to_Lat(fEastings, fNorthings, 6377340.189, 6356034.447, 200000, 250000, 1.000035, 53.5, -8#)
      lng1 = E_N_to_Long(fEastings, fNorthings, 6377340.189, 6356034.447, 200000, 250000, 1.000035, 53.5, -8#)

      'Get cartesian coordinates
      X1 = Lat_Long_H_to_X(lat1, lng1, height, 6377340.189, 6356034.447)
      Y1 = Lat_Long_H_to_Y(lat1, lng1, height, 6377340.189, 6356034.447)
      Z1 = Lat_H_to_Z(lat1, height, 6377340.189, 6356034.447)

      'Offset coordinates to real coordinates
      x2 = Helmert_X(X1, Y1, Z1, 482.53, 0.214, 0.631, 8.15)
      Y2 = Helmert_Y(X1, Y1, Z1, -130.596, 1.042, 0.631, 8.15)
      Z2 = Helmert_Z(X1, Y1, Z1, 564.557, 1.042, 0.214, 8.15)

      'Convert back to latitude, longitude
      fLatitude = XYZ_to_Lat(x2, Y2, Z2, 6378137#, 6356752.313)
      fLongitude = XYZ_to_Long(x2, Y2)
      
      'Initialise coord data
      Call pnt.protInit(LatLng, fLongitude, fLatitude)
      Call pnt.protInit(NGR_Ireland, fEastings, fNorthings)
    Case Else
      Err.Raise 1, "", "No type for this GridRef"
  End Select
  Set CreateFromProjection = pnt
End Function

'Create a point from EN (425342,432322), national grid reference (6 char or 10char) (SK00000000),
'@param {String}  sReference  - String reference to be parsed to coordinates
'@param {ENGRType?} iNGRType  - BritishNationalGrid or IrishNationalGrid depending on coord type
'@returns {gisPoint} A point object
Public Function CreateFromString(ByVal sReference As String, Optional ByVal iNGRType As ENGRType) As gisPoint
  Select Case iNGRType
    Case BritishNationalGrid
      sGridReference = Trim(sGridReference)
  
      Dim oMatch As Object
      Static rxDigString As Object: If rxDigString Is Nothing Then Set rxDigString = regex("^(\d+)\s*[, ;-]\s*(\d+)$")
      If rxDigString.Test(sGridReference) Then
        Set oMatch = rxDigString.Execute(sGridReference)
        Set CreateFromString = CreateFromProjection(oMatch(0).submatches(0), oMatch(0).submatches(1), BritishNationalGrid)
        Exit Function
      End If
      
      Static rxGridRef10 As Object: If rxGridRef10 Is Nothing Then Set rxGridRef10 = regex("^([A-Z]{2})\s*([0-9]{5})\s*([0-9]{5})$", "i")
      Static rxGridRef6 As Object: If rxGridRef6 Is Nothing Then Set rxGridRef6 = regex("^([A-Z]{2})\s*([0-9]{3})\s*([0-9]{3})$", "i")
      Dim iDig As Long: iDig = 0
      Select Case True
        Case rxGridRef10.Test(sGridReference)
          iDig = 10
        Case rxGridRef6.Test(sGridReference)
          iDig = 6
      End Select
      
      If iDig > 0 Then
        'Get matched regex
        If iDig = 10 Then
          Set oMatch = rxGridRef10.Execute(sGridReference)
        Else
          Set oMatch = rxGridRef6.Execute(sGridReference)
        End If
        
        'First 2 letters --> e100km and n100km
        Dim sChars As String: sChars = oMatch(0).submatches(0)
        
        'get numeric values of letter references, mapping A->0, B->1, C->2, etc:
        Dim i1 As Long: i1 = Asc(UCase(mid(sChars, 1, 1))) - Asc("A")
        Dim i2 As Long: i2 = Asc(UCase(mid(sChars, 2, 1))) - Asc("A")
        
        'shuffle down letters after 'I' since 'I' is not used in grid:
        If i1 > 7 Then i1 = i1 - 1
        If i2 > 7 Then i2 = i2 - 1
        
        'sanity check
        If (i1 < 8 Or i1 > 18) Then Error.Raise 1, "gisPoint::CreateFromString", "Invalid grid reference """ & sGridReference & """"
        
        'convert grid letters into 100km-square indexes from false origin (grid square SV):
        Dim e100km As Double: e100km = ((i1 - 2) Mod 5) * 5 + (i2 Mod 5)
        Dim n100km As Double: n100km = (19 - floor(i1 / 5) * 5) - floor(i2 / 5)
        
        'Find easting and northing coords
        Dim fEasting As Long, fNorthing As Long
        If iDig = 10 Then
          fEasting = e100km & oMatch(0).submatches(1)
          fNorthing = n100km & oMatch(0).submatches(2)
        Else
          fEasting = CLng(e100km & oMatch(0).submatches(1) & "00")
          fNorthing = CLng(n100km & oMatch(0).submatches(2) & "00")
        End If
        
        'Create point
        Set CreateFromString = CreateFromProjection(fEasting, fNorthing, BritishNationalGrid)
      Else
        Err.Raise 1, "gisOSGB36::CreateFromString", "Invalid grid reference """ & sGridReference & """"
      End If
    Case IrishNationalGrid
      'TODO:
    Case Else
      Err.Raise 1, "", "No type for this GridRef"
  End Select
End Function


'Obtain the latitude coord
Public Property Get Latitude() As Double
  Latitude = pY(LatLng)
End Property

'Obtain the longitude coord
Public Property Get Longitude() As Double
  Longitude = pX(LatLng)
End Property

'Obtain the eastings coord
'@param {ENGRType} BritishNationalGrid or IrishNationalGrid
Public Property Get Eastings(Optional ByVal iNGRType As ENGRType = BritishNationalGrid) As Double
  Select Case iNGRType
    Case BritishNationalGrid
      If Not pB(NGR_UK) Then Call initNGRUK
      Eastings = pX(NGR_UK)
    Case IrishNationalGrid
      If Not pB(NGR_Ireland) Then Call initNGRIreland
      Eastings = pX(NGR_Ireland)
  End Select
End Property

'Obtain the northings coord
'@param {ENGRType} BritishNationalGrid or IrishNationalGrid
Public Property Get Northings(Optional ByVal iNGRType As ENGRType = BritishNationalGrid) As Double
  Select Case iNGRType
    Case BritishNationalGrid
      If Not pB(NGR_UK) Then Call initNGRUK
      Northings = pY(NGR_UK)
    Case IrishNationalGrid
      If Not pB(NGR_Ireland) Then Call initNGRIreland
      Northings = pY(NGR_Ireland)
  End Select
End Property


'*********************
'* PROTECTED METHODS *
'*********************

'Initialise coord x and y of a given type
'@param {ECoordType} iType - Coordinate type to update
'@param {Double}     fX    - X coordinate
'@param {Double}     fY    - Y coordinate
Friend Sub protInit(ByVal iType As ECoordType, ByVal fX As Double, ByVal fY As Double)
  pB(iType) = True
  pX(iType) = fX
  pY(iType) = fY
End Sub


'********************
'* HELPER FUNCTIONS *
'********************

'Convert NGRType to CoordType
'@param {ENGRType} National grid type
'@returns {ECoordType} Coordinate type
Private Function NGRToCoordType(ByVal iNGRType As ENGRType) As ECoordType
  Select Case iNGRType
    Case BritishNationalGrid: NGRToCoordType = ECoordType.NGR_UK
    Case IrishNationalGrid: NGRToCoordType = ECoordType.NGR_Ireland
  End Select
End Function

'Initialise Eastings and Northings of NGR_UK from Latitude Longitude
Private Sub initNGRUK()
  If isInGreatBritain(pY(LatLng), pX(LatLng)) Then
    '
    Dim height As Double: height = 0
    Dim X1 As Double: X1 = Lat_Long_H_to_X(pY(LatLng), pX(LatLng), height, 6378137#, 6356752.313)
    Dim Y1 As Double: Y1 = Lat_Long_H_to_Y(pY(LatLng), pX(LatLng), height, 6378137#, 6356752.313)
    Dim Z1 As Double: Z1 = Lat_H_to_Z(pY(LatLng), height, 6378137#, 6356752.313)

    Dim x2 As Double: x2 = Helmert_X(X1, Y1, Z1, -446.448, -0.247, -0.8421, 20.4894)
    Dim Y2 As Double: Y2 = Helmert_Y(X1, Y1, Z1, 125.157, -0.1502, -0.8421, 20.4894)
    Dim Z2 As Double: Z2 = Helmert_Z(X1, Y1, Z1, -542.06, -0.1502, -0.247, 20.4894)

    Dim latitude2 As Double: latitude2 = XYZ_to_Lat(x2, Y2, Z2, 6377563.396, 6356256.91)
    Dim longitude2 As Double: longitude2 = XYZ_to_Long(x2, Y2)

    Dim fEastings As Double: fEastings = Lat_Long_to_East(latitude2, longitude2, 6377563.396, 6356256.91, 400000, 0.999601272, 49#, -2#)
    Dim fNorthing As Double: fNorthing = Lat_Long_to_North(latitude2, longitude2, 6377563.396, 6356256.91, 400000, -100000, 0.999601272, 49#, -2#)

    pB(NGR_UK) = True
    pX(NGR_UK) = fEastings
    pY(NGR_UK) = fNorthing
  Else
    Err.Raise 1, "", "Point not in great britain"
  End If
End Sub

'Initialise Eastings and Northings of NGR_Ireland from Latitude Longitude
Private Sub initNGRIreland()
  If isInIreland(pY(LatLng), pX(LatLng)) Then
    '
    Dim height As Double: height = 0
    Dim X1 As Double: X1 = Lat_Long_H_to_X(pY(LatLng), pX(LatLng), height, 6378137#, 6356752.313)
    Dim Y1 As Double: Y1 = Lat_Long_H_to_Y(pY(LatLng), pX(LatLng), height, 6378137#, 6356752.313)
    Dim Z1 As Double: Z1 = Lat_H_to_Z(pY(LatLng), height, 6378137#, 6356752.313)

    Dim x2 As Double: x2 = Helmert_X(X1, Y1, Z1, -482.53, -0.214, -0.631, -8.15)
    Dim Y2 As Double: Y2 = Helmert_Y(X1, Y1, Z1, 130.596, -1.042, -0.631, -8.15)
    Dim Z2 As Double: Z2 = Helmert_Z(X1, Y1, Z1, -564.557, -1.042, -0.214, -8.15)

    Dim latitude2 As Double: latitude2 = XYZ_to_Lat(x2, Y2, Z2, 6377340.189, 6356034.447)
    Dim longitude2 As Double: longitude2 = XYZ_to_Long(x2, Y2)

    Dim fEastings As Double: fEastings = Lat_Long_to_East(latitude2, longitude2, 6377340.189, 6356034.447, 200000, 1.000035, 53.5, -8#)
    Dim fNorthing As Double: fNorthing = Lat_Long_to_North(latitude2, longitude2, 6377340.189, 6356034.447, 200000, 250000, 1.000035, 53.5, -8#)

    pB(NGR_Ireland) = True
    pX(NGR_Ireland) = fEastings
    pY(NGR_Ireland) = fNorthing
  Else
    Err.Raise 1, "", "Point not in ireland"
  End If
End Sub


'Check if the latitude and longitude coord are in the bounding box of the UK
'@param {Double} fLatitude  - Latitude coord
'@param {Double} fLongitude - Longitude coord
'@returns {Boolean} True if in great britain, false otherwise
Private Function isInGreatBritain(ByVal fLatitude As Double, ByVal fLongitude As Double) As Boolean
  isInGreatBritain = fLatitude > 49 And fLatitude < 62 And fLongitude > -9.5 And fLongitude < 2.3
End Function

'Check if the latitude and longitude coord are in the bounding box of the Ireland
'@param {Double}
'@param {Double}
'@returns {Boolean} True if in ireland, false otherwise
Private Function isInIreland(ByVal fLatitude As Double, ByVal fLongitude As Double) As Boolean
  isInIreland = fLatitude > 51.2 And fLatitude < 55.73 And fLongitude > -12.2 And fLongitude < -4.8
End Function

'Computed Helmert transformed X coordinate.
'@param {Double} X     - Cartesian X coord in meters
'@param {Double} Y     - Cartesian Y coord in meters
'@param {Double} Z     - Cartesian Z coord in meters
'@param {Double} DX    - Cartesian X translation in meters
'@param {Double} Y_Rot - Y rotation in seconds of arc
'@param {Double} Z_Rot - Z rotation in seconds of arc
'@param {Double} s     - Scale in ppm
'@returns {Double} Helmert transformed X coordinate
Private Function Helmert_X(ByVal x As Double, ByVal y As Double, ByVal z As Double, ByVal DX As Double, ByVal Y_Rot As Double, ByVal Z_Rot As Double, ByVal s As Double)
  'Convert rotations to radians and ppm scale to a factor
  Dim sfactor As Double: sfactor = s * 0.000001
  Dim RadY_Rot As Double: RadY_Rot = (Y_Rot / 3600) * (PI / 180)
  Dim RadZ_Rot As Double: RadZ_Rot = (Z_Rot / 3600) * (PI / 180)

  'Compute transformed X coord
  Helmert_X = x + (x * sfactor) - (y * RadZ_Rot) + (z * RadY_Rot) + DX
End Function

'Computed Helmert transformed Y coordinate.
'@param {Double} X     - Cartesian X coord in meters
'@param {Double} Y     - Cartesian Y coord in meters
'@param {Double} Z     - Cartesian Z coord in meters
'@param {Double} DY    - Cartesian Y translation in meters
'@param {Double} X_Rot - X rotation in seconds of arc
'@param {Double} Z_Rot - Z rotation in seconds of arc
'@param {Double} s     - Scale in ppm
'@returns {Double} Helmert transformed Y coordinate
Private Function Helmert_Y(ByVal x As Double, ByVal y As Double, ByVal z As Double, ByVal DY As Double, ByVal X_Rot As Double, ByVal Z_Rot As Double, ByVal s As Double)
  'Convert rotations to radians and ppm scale to a factor
  Dim sfactor As Double: sfactor = s * 0.000001
  Dim RadX_Rot As Double: RadX_Rot = (X_Rot / 3600) * (PI / 180)
  Dim RadZ_Rot As Double: RadZ_Rot = (Z_Rot / 3600) * (PI / 180)

  'Compute transformed Y coord
  Helmert_Y = (x * RadZ_Rot) + y + (y * sfactor) - (z * RadX_Rot) + DY
End Function

'Computed Helmert transformed Z coordinate.
'@param {Double} X     - Cartesian X coord in meters
'@param {Double} Y     - Cartesian Y coord in meters
'@param {Double} Z     - Cartesian Z coord in meters
'@param {Double} DZ    - Cartesian Z translation in meters
'@param {Double} X_Rot - X rotation in seconds of arc
'@param {Double} Y_Rot - Y rotation in seconds of arc
'@param {Double} s     - Scale in ppm
'@returns {Double} Helmert transformed Z coordinate
Private Function Helmert_Z(x, y, z, DZ, X_Rot, Y_Rot, s)
  'Convert rotations to radians and ppm scale to a factor
  Dim sfactor As Double: sfactor = s * 0.000001
  Dim RadX_Rot As Double: RadX_Rot = (X_Rot / 3600) * (PI / 180)
  Dim RadY_Rot As Double: RadY_Rot = (Y_Rot / 3600) * (PI / 180)

  'Compute transformed Z coord
  Helmert_Z = (-1 * x * RadY_Rot) + (y * RadX_Rot) + z + (z * sfactor) + DZ
End Function

'Convert XYZ to Latitude (PHI) in Dec Degrees.
'@param {Double} X     - Cartesian X coord in meters
'@param {Double} Y     - Cartesian Y coord in meters
'@param {Double} Z     - Cartesian Z coord in meters
'@param {Double} a     - Ellipsoid semi-major axis in meters
'@param {Double} b     - Ellipsoid semi-minor axis in meters
'@returns {Double} Latitude in decimal (TODO:degrees or radians?)
'@dependencies [Iterate_XYZ_to_Lat]
Private Function XYZ_to_Lat(ByVal x As Double, ByVal y As Double, ByVal z As Double, ByVal a As Double, ByVal b As Double) As Double
  Dim RootXYSqr As Double: RootXYSqr = Sqr((x ^ 2) + (y ^ 2))
  Dim e2 As Double: e2 = ((a ^ 2) - (b ^ 2)) / (a ^ 2)
  Dim PHI1 As Double: PHI1 = Atn(z / (RootXYSqr * (1 - e2)))
  Dim phi As Double: phi = Iterate_XYZ_to_Lat(a, e2, PHI1, z, RootXYSqr)
  
  XYZ_to_Lat = phi * (180 / PI)
End Function

'Iteratively computes Latitude (PHI).
'@param {Double} a          - Elipsoid semi-major axis in meters
'@param {Double} e2         - eta squared
'@param {Double} PHI1       - estimated value for latitude in radians
'@param {Double} Z          - Cartesian coordinate in meters
'@param {Double} RootXYSqr  - Computed from X and Y in meters
'@returns {Double} Latitude in decimal (TODO: degrees or radians?)
Private Function Iterate_XYZ_to_Lat(ByVal a As Double, ByVal e2 As Double, ByVal PHI1 As Double, ByVal z As Double, ByVal RootXYSqr As Double) As Double
  Dim v As Double: v = a / (Sqr(1 - (e2 * ((Sin(PHI1)) ^ 2))))
  Dim PHI2 As Double: PHI2 = Atn((z + (e2 * v * (Sin(PHI1)))) / RootXYSqr)
  
  Do While Abs(PHI1 - PHI2) > 0.000000001
    PHI1 = PHI2
    v = a / (Sqr(1 - (e2 * ((Sin(PHI1)) ^ 2))))
    PHI2 = Atn((z + (e2 * v * (Sin(PHI1)))) / RootXYSqr)
  Loop

  Iterate_XYZ_to_Lat = PHI2
End Function

'Convert XYZ to Longitude (LAM) in Decimal Degrees.
'@param {Double} X     - Cartesian X coord in meters
'@param {Double} Y     - Cartesian Y coord in meters
'@returns {Double} Longitude in decimal degrees
Private Function XYZ_to_Long(ByVal x As Double, ByVal y As Double) As Double
  'tailor the output to fit the equatorial quadrant as determined by the signs of X and Y
  If x >= 0 Then 'longitude is in the W90 thru 0 to E90 hemisphere
    XYZ_to_Long = (Atn(y / x)) * (180 / PI)
  End If
  
  If x < 0 And y >= 0 Then 'longitude is in the E90 to E180 quadrant
    XYZ_to_Long = ((Atn(y / x)) * (180 / PI)) + 180
  End If
  
  If x < 0 And y < 0 Then 'longitude is in the E180 to W90 quadrant
    XYZ_to_Long = ((Atn(y / x)) * (180 / PI)) - 180
  End If
End Function

'Convert XYZ to Ellipsoidal Height (H).
'@param {Double} X     - Cartesian X coord in meters
'@param {Double} Y     - Cartesian Y coord in meters
'@param {Double} Z     - Cartesian Z coord in meters
'@param {Double} a     - Ellipsoid semi-major axis in meters
'@param {Double} b     - Ellipsoid semi-minor axis in meters
'@returns {Double}
'@dependencies [XYZ_to_Lat()]
Private Function XYZ_to_H(ByVal x As Double, ByVal y As Double, ByVal z As Double, ByVal a As Double, ByVal b As Double) As Double
  'Compute PHI (Dec Degrees) first
  Dim phi As Double: phi = XYZ_to_Lat(x, y, z, a, b)

  'Convert PHI radians
  Dim RadPHI As Double: RadPHI = phi * (PI / 180)

  'Compute Elipsoidal Height H
  Dim RootXYSqr As Double: RootXYSqr = Sqr((x ^ 2) + (y ^ 2))
  Dim e2 As Double: e2 = ((a ^ 2) - (b ^ 2)) / (a ^ 2)
  Dim v As Double: v = a / (Sqr(1 - (e2 * ((Sin(RadPHI)) ^ 2))))
  Dim h As Double: h = (RootXYSqr / Cos(RadPHI)) - v

  XYZ_to_H = h
End Function

'Convert geodetic coords lat (PHI), long (LAM) and height (H) to cartesian X coordinate.
'@param {Double} PHI   - Latitude in decimal degrees
'@param {Double} LAM   - Longitude in decimal degrees
'@param {Double} H     - Ellipsoid height in meters
'@param {Double} a     - Ellipsoid semi-major axis in meters
'@param {Double} b     - Ellipsoid semi-minor axis in meters
'@returns {Double}     - Cartesian X coord in meters
Private Function Lat_Long_H_to_X(ByVal phi As Double, ByVal LAM As Double, ByVal h As Double, ByVal a As Double, ByVal b As Double) As Double
  'Convert angle measures to radians
  Dim RadPHI As Double: RadPHI = phi * (PI / 180)
  Dim RadLAM As Double: RadLAM = LAM * (PI / 180)

  'Compute eccentricity squared and nu
  Dim e2 As Double: e2 = ((a ^ 2) - (b ^ 2)) / (a ^ 2)
  Dim v As Double: v = a / (Sqr(1 - (e2 * ((Sin(RadPHI)) ^ 2))))

  'Compute X
  Lat_Long_H_to_X = (v + h) * (Cos(RadPHI)) * (Cos(RadLAM))
End Function

'Convert geodetic coords lat (PHI), long (LAM) and height (H) to cartesian Y coordinate.
'@param {Double} PHI   - Latitude in decimal degrees
'@param {Double} LAM   - Longitude in decimal degrees
'@param {Double} H     - Ellipsoid height in meters
'@param {Double} a     - Ellipsoid semi-major axis in meters
'@param {Double} b     - Ellipsoid semi-minor axis in meters
'@returns {Double}     - Cartesian Y coord in meters
Private Function Lat_Long_H_to_Y(ByVal phi As Double, ByVal LAM As Double, ByVal h As Double, ByVal a As Double, ByVal b As Double) As Double
  'Convert angle measures to radians
  Dim RadPHI As Double: RadPHI = phi * (PI / 180)
  Dim RadLAM As Double: RadLAM = LAM * (PI / 180)

  'Compute eccentricity squared and nu
  Dim e2 As Double: e2 = ((a ^ 2) - (b ^ 2)) / (a ^ 2)
  Dim v As Double: v = a / (Sqr(1 - (e2 * ((Sin(RadPHI)) ^ 2))))

  'Compute Y
  Lat_Long_H_to_Y = (v + h) * (Cos(RadPHI)) * (Sin(RadLAM))
End Function

'Convert geodetic coords lat (PHI) and height (H) to cartesian Z coordinate.
'@param {Double} PHI   - Latitude in decimal degrees
'@param {Double} H     - Ellipsoid height in meters
'@param {Double} a     - Ellipsoid semi-major axis in meters
'@param {Double} b     - Ellipsoid semi-minor axis in meters
'@returns {Double}     - Cartesian Z coord in meters
Private Function Lat_H_to_Z(ByVal phi As Double, ByVal h As Double, ByVal a As Double, ByVal b As Double) As Double
  'Convert angle measures to radians
  Dim RadPHI As Double: RadPHI = phi * (PI / 180)

  'Compute eccentricity squared and nu
  Dim e2 As Double: e2 = ((a ^ 2) - (b ^ 2)) / (a ^ 2)
  Dim v As Double: v = a / (Sqr(1 - (e2 * ((Sin(RadPHI)) ^ 2))))

  'Compute Z
  Lat_H_to_Z = ((v * (1 - e2)) + h) * (Sin(RadPHI))
End Function

'Project Latitude and longitude to Transverse Mercator Eastings.
'@param {Double} PHI   - Latitude in decimal degrees
'@param {Double} LAM   - Longitude in decimal degrees
'@param {Double} a     - Ellipsoid semi-major axis in meters
'@param {Double} b     - Ellipsoid semi-minor axis in meters
'@param {Double} e0    - Eastings of false origin in meters
'@param {Double} f0    - Central meridian scale factor
'@param {Double} PHI0  - Latitude of false origin in decimal degrees
'@param {Double} LAM0  - Longitude of false origin in decimal degrees
Private Function Lat_Long_to_East(ByVal phi As Double, ByVal LAM As Double, ByVal a As Double, ByVal b As Double, ByVal E0 As Double, ByVal F0 As Double, ByVal phi0 As Double, ByVal LAM0 As Double) As Double
  'Convert angle measures to radians
  Dim RadPHI As Double: RadPHI = phi * (PI / 180)
  Dim RadLAM As Double: RadLAM = LAM * (PI / 180)
  Dim RadPHI0 As Double: RadPHI0 = phi0 * (PI / 180)
  Dim RadLAM0 As Double: RadLAM0 = LAM0 * (PI / 180)

  Dim af0 As Double: af0 = a * F0
  Dim bf0 As Double: bf0 = b * F0
  Dim e2 As Double: e2 = ((af0 ^ 2) - (bf0 ^ 2)) / (af0 ^ 2)
  Dim n As Double: n = (af0 - bf0) / (af0 + bf0)
  Dim nu As Double: nu = af0 / (Sqr(1 - (e2 * ((Sin(RadPHI)) ^ 2))))
  Dim rho As Double: rho = (nu * (1 - e2)) / (1 - (e2 * (Sin(RadPHI)) ^ 2))
  Dim eta2 As Double: eta2 = (nu / rho) - 1
  Dim p As Double: p = RadLAM - RadLAM0

  Dim IV As Double: IV = nu * (Cos(RadPHI))
  Dim v As Double: v = (nu / 6) * ((Cos(RadPHI)) ^ 3) * ((nu / rho) - ((Tan(RadPHI) ^ 2)))
  Dim VI As Double: VI = (nu / 120) * ((Cos(RadPHI)) ^ 5) * (5 - (18 * ((Tan(RadPHI)) ^ 2)) + ((Tan(RadPHI)) ^ 4) + (14 * eta2) - (58 * ((Tan(RadPHI)) ^ 2) * eta2))

  Lat_Long_to_East = E0 + (p * IV) + ((p ^ 3) * v) + ((p ^ 5) * VI)
End Function

'Project Latitude and longitude to Transverse Mercator Northings.
'@param {Double} PHI   - Latitude in decimal degrees
'@param {Double} LAM   - Longitude in decimal degrees
'@param {Double} a     - Ellipsoid semi-major axis in meters
'@param {Double} b     - Ellipsoid semi-minor axis in meters
'@param {Double} e0    - Eastings of false origin in meters
'@param {Double} n0    - Northings of false origin in meters
'@param {Double} f0    - Central meridian scale factor
'@param {Double} PHI0  - Latitude of false origin in decimal degrees
'@param {Double} LAM0  - Longitude of false origin in decimal degrees
'@dependencies [Marc()]
Private Function Lat_Long_to_North(ByVal phi As Double, ByVal LAM As Double, ByVal a As Double, ByVal b As Double, ByVal E0 As Double, ByVal N0 As Double, ByVal F0 As Double, ByVal phi0 As Double, ByVal LAM0 As Double) As Double
  'Convert angle measures to radians
  Dim RadPHI As Double: RadPHI = phi * (PI / 180)
  Dim RadLAM As Double: RadLAM = LAM * (PI / 180)
  Dim RadPHI0 As Double: RadPHI0 = phi0 * (PI / 180)
  Dim RadLAM0 As Double: RadLAM0 = LAM0 * (PI / 180)

  Dim af0 As Double: af0 = a * F0
  Dim bf0 As Double: bf0 = b * F0
  Dim e2 As Double: e2 = ((af0 ^ 2) - (bf0 ^ 2)) / (af0 ^ 2)
  Dim n As Double: n = (af0 - bf0) / (af0 + bf0)
  Dim nu As Double: nu = af0 / (Sqr(1 - (e2 * ((Sin(RadPHI)) ^ 2))))
  Dim rho As Double: rho = (nu * (1 - e2)) / (1 - (e2 * (Sin(RadPHI)) ^ 2))
  Dim eta2 As Double: eta2 = (nu / rho) - 1
  Dim p As Double: p = RadLAM - RadLAM0
  Dim M As Double: M = Marc(bf0, n, RadPHI0, RadPHI)

  Dim i As Double: i = M + N0
  Dim II As Double: II = (nu / 2) * (Sin(RadPHI)) * (Cos(RadPHI))
  Dim III As Double: III = ((nu / 24) * (Sin(RadPHI)) * ((Cos(RadPHI)) ^ 3)) * (5 - ((Tan(RadPHI)) ^ 2) + (9 * eta2))
  Dim IIIA As Double: IIIA = ((nu / 720) * (Sin(RadPHI)) * ((Cos(RadPHI)) ^ 5)) * (61 - (58 * ((Tan(RadPHI)) ^ 2)) + ((Tan(RadPHI)) ^ 4))

  Lat_Long_to_North = i + ((p ^ 2) * II) + ((p ^ 4) * III) + ((p ^ 6) * IIIA)
End Function

'Un-project Transverse Mercator eastings and northings back to latitude.
'@param {Double} East  - Eastings in meters
'@param {Double} North - Northings in meters
'@param {Double} a     - Ellipsoid semi-major axis in meters
'@param {Double} b     - Ellipsoid semi-minor axis in meters
'@param {Double} e0    - Eastings of false origin in meters
'@param {Double} n0    - Northings of false origin in meters
'@param {Double} f0    - Central meridian scale factor
'@param {Double} PHI0  - Latitude of false origin in decimal degrees
'@param {Double} LAM0  - Longitude of false origin in decimal degrees
'@returns {Double}     - Latitude in decimal degrees
'@dependencies [Marc(),InitialLat()]
Private Function E_N_to_Lat(ByVal East As Double, ByVal North As Double, ByVal a As Double, ByVal b As Double, ByVal E0 As Double, ByVal N0 As Double, ByVal F0 As Double, ByVal phi0 As Double, ByVal LAM0 As Double) As Double
  'Convert angle measures to radians
  Dim RadPHI0 As Double: RadPHI0 = phi0 * (PI / 180)
  Dim RadLAM0 As Double: RadLAM0 = LAM0 * (PI / 180)

  'Compute af0, bf0, e squared (e2), n and Et
  Dim af0 As Double: af0 = a * F0
  Dim bf0 As Double: bf0 = b * F0
  Dim e2 As Double: e2 = ((af0 ^ 2) - (bf0 ^ 2)) / (af0 ^ 2)
  Dim n As Double: n = (af0 - bf0) / (af0 + bf0)
  Dim Et As Double: Et = East - E0

  'Compute initial value for latitude (PHI) in radians
  Dim PHId As Double: PHId = InitialLat(North, N0, af0, RadPHI0, n, bf0)

  'Compute nu, rho and eta2 using value for PHId
  Dim nu As Double: nu = af0 / (Sqr(1 - (e2 * ((Sin(PHId)) ^ 2))))
  Dim rho As Double: rho = (nu * (1 - e2)) / (1 - (e2 * (Sin(PHId)) ^ 2))
  Dim eta2 As Double: eta2 = (nu / rho) - 1

  'Compute Latitude
  Dim VII As Double: VII = (Tan(PHId)) / (2 * rho * nu)
  Dim VIII As Double: VIII = ((Tan(PHId)) / (24 * rho * (nu ^ 3))) * (5 + (3 * ((Tan(PHId)) ^ 2)) + eta2 - (9 * eta2 * ((Tan(PHId)) ^ 2)))
  Dim IX As Double: IX = ((Tan(PHId)) / (720 * rho * (nu ^ 5))) * (61 + (90 * ((Tan(PHId)) ^ 2)) + (45 * ((Tan(PHId)) ^ 4)))

  E_N_to_Lat = (180 / PI) * (PHId - ((Et ^ 2) * VII) + ((Et ^ 4) * VIII) - ((Et ^ 6) * IX))
End Function

'Un-project Transverse Mercator eastings and northings back to longitude.
'@param {Double} East  - Eastings in meters
'@param {Double} North - Northings in meters
'@param {Double} a     - Ellipsoid semi-major axis in meters
'@param {Double} b     - Ellipsoid semi-minor axis in meters
'@param {Double} e0    - Eastings of false origin in meters
'@param {Double} n0    - Northings of false origin in meters
'@param {Double} f0    - Central meridian scale factor
'@param {Double} PHI0  - Latitude of false origin in decimal degrees
'@param {Double} LAM0  - Longitude of false origin in decimal degrees
'@returns {Double}     - Longitude in decimal degrees
'@dependencies [Marc(),InitialLat()]
Private Function E_N_to_Long(ByVal East As Double, ByVal North As Double, ByVal a As Double, ByVal b As Double, ByVal E0 As Double, ByVal N0 As Double, ByVal F0 As Double, ByVal phi0 As Double, ByVal LAM0 As Double) As Double
  'Convert angle measures to radians
  Dim RadPHI0 As Double: RadPHI0 = phi0 * (PI / 180)
  Dim RadLAM0 As Double: RadLAM0 = LAM0 * (PI / 180)

  'Compute af0, bf0, e squared (e2), n and Et
  Dim af0 As Double: af0 = a * F0
  Dim bf0 As Double: bf0 = b * F0
  Dim e2 As Double: e2 = ((af0 ^ 2) - (bf0 ^ 2)) / (af0 ^ 2)
  Dim n As Double: n = (af0 - bf0) / (af0 + bf0)
  Dim Et As Double: Et = East - E0

  'Compute initial value for latitude (PHI) in radians
  Dim PHId As Double: PHId = InitialLat(North, N0, af0, RadPHI0, n, bf0)

  'Compute nu, rho and eta2 using value for PHId
  Dim nu As Double: nu = af0 / (Sqr(1 - (e2 * ((Sin(PHId)) ^ 2))))
  Dim rho As Double: rho = (nu * (1 - e2)) / (1 - (e2 * (Sin(PHId)) ^ 2))
  Dim eta2 As Double: eta2 = (nu / rho) - 1

  'Compute Longitude
  Dim x As Double: x = ((Cos(PHId)) ^ -1) / nu
  Dim XI As Double: XI = (((Cos(PHId)) ^ -1) / (6 * (nu ^ 3))) * ((nu / rho) + (2 * ((Tan(PHId)) ^ 2)))
  Dim XII As Double: XII = (((Cos(PHId)) ^ -1) / (120 * (nu ^ 5))) * (5 + (28 * ((Tan(PHId)) ^ 2)) + (24 * ((Tan(PHId)) ^ 4)))
  Dim XIIA As Double: XIIA = (((Cos(PHId)) ^ -1) / (5040 * (nu ^ 7))) * (61 + (662 * ((Tan(PHId)) ^ 2)) + (1320 * ((Tan(PHId)) ^ 4)) + (720 * ((Tan(PHId)) ^ 6)))

  E_N_to_Long = (180 / PI) * (RadLAM0 + (Et * x) - ((Et ^ 3) * XI) + ((Et ^ 5) * XII) - ((Et ^ 7) * XIIA))
End Function

'Compute initial value for Latitude (PHI) IN RADIANS.
'@param {Double} North - Northings of point in meters
'@param {Double} n0    - Northings of false origin in meters
'@param {Double} afo   - semi major axis multiplied by central meridian scale factor in meters
'@param {Double} PHI0  - Latitude of false origin in radians
'@param {Double} n     - (computed from a, b and f0)
'@param {Double} bfo   - ellipsoid semi major axis multiplied by central meridian scale factor in meters
'@returns {Double} initial value for Latitude in radians
'@dependencies [Marc()]
Private Function InitialLat(ByVal North As Double, ByVal N0 As Double, ByVal afo As Double, ByVal phi0 As Double, ByVal n As Double, ByVal bfo As Double) As Double
  'First PHI value (PHI1)
  Dim PHI1 As Double: PHI1 = ((North - N0) / afo) + phi0

  'Calculate M
  Dim M As Double: M = Marc(bfo, n, phi0, PHI1)

  'Calculate new PHI value (PHI2)
  Dim PHI2 As Double: PHI2 = ((North - N0 - M) / afo) + PHI1

  'Iterate to get final value for InitialLat
  Do While Abs(North - N0 - M) > 0.00001
    PHI2 = ((North - N0 - M) / afo) + PHI1
    M = Marc(bfo, n, phi0, PHI2)
    PHI1 = PHI2
  Loop

  InitialLat = PHI2
End Function

'Compute meridional arc.
'@param {Double} bfo   - ellipsoid semi major axis multiplied by central meridian scale factor in meters
'@param {Double} n     - (computed from a, b and f0)
'@param {Double} PHI0  - Latitude of false origin in radians
'@param {Double} PHI   - Final latitude of point in radians
'@returns {Double} meridional arc
Private Function Marc(ByVal bf0 As Double, ByVal n As Double, ByVal phi0 As Double, ByVal phi As Double) As Double
  Marc = bf0 * (((1 + n + ((5 / 4) * (n ^ 2)) + ((5 / 4) * (n ^ 3))) * (phi - phi0)) _
  - (((3 * n) + (3 * (n ^ 2)) + ((21 / 8) * (n ^ 3))) * (Sin(phi - phi0)) * (Cos(phi + phi0))) _
  + ((((15 / 8) * (n ^ 2)) + ((15 / 8) * (n ^ 3))) * (Sin(2 * (phi - phi0))) * (Cos(2 * (phi + phi0)))) _
  - (((35 / 24) * (n ^ 3)) * (Sin(3 * (phi - phi0))) * (Cos(3 * (phi + phi0)))))
End Function

'Compute convergence (in decimal degrees) from latitude and longitude
'@param {Double} PHI   - Latitude in decimal degrees
'@param {Double} LAM   - Longitude in decimal degrees
'@param {Double} LAM0  - Longitude of false origin in decimal degrees
'@param {Double} a     - Ellipsoid semi-major axis in meters
'@param {Double} b     - Ellipsoid semi-minor axis in meters
'@param {Double} f0    - Central meridian scale factor
'@returns {Double}     - Convergence in decimal degrees
Private Function Lat_Long_to_C(ByVal phi As Double, ByVal LAM As Double, ByVal LAM0 As Double, ByVal a As Double, ByVal b As Double, ByVal F0 As Double) As Double
  'Convert angle measures to radians
  Dim RadPHI As Double: RadPHI = phi * (PI / 180)
  Dim RadLAM As Double: RadLAM = LAM * (PI / 180)
  Dim RadLAM0 As Double: RadLAM0 = LAM0 * (PI / 180)
      
  'Compute af0, bf0 and e squared (e2)
  Dim af0 As Double: af0 = a * F0
  Dim bf0 As Double: bf0 = b * F0
  Dim e2 As Double: e2 = ((af0 ^ 2) - (bf0 ^ 2)) / (af0 ^ 2)

  'Compute nu, rho, eta2 and p
  Dim nu As Double: nu = af0 / (Sqr(1 - (e2 * ((Sin(RadPHI)) ^ 2))))
  Dim rho As Double: rho = (nu * (1 - e2)) / (1 - (e2 * (Sin(RadPHI)) ^ 2))
  Dim eta2 As Double: eta2 = (nu / rho) - 1
  Dim p As Double: p = RadLAM - RadLAM0

  'Compute Convergence
  Dim XIII As Double: XIII = Sin(RadPHI)
  Dim XIV As Double: XIV = ((Sin(RadPHI) * ((Cos(RadPHI)) ^ 2)) / 3) * (1 + (3 * eta2) + (2 * (eta2 ^ 2)))
  Dim XV As Double: XV = ((Sin(RadPHI) * ((Cos(RadPHI)) ^ 4)) / 15) * (2 - ((Tan(RadPHI)) ^ 2))

  Lat_Long_to_C = (180 / PI) * ((p * XIII) + ((p ^ 3) * XIV) + ((p ^ 5) * XV))
End Function


'Compute convergence (in decimal degrees) from easting and northing
'@param {Double} East  - Eastings in meters
'@param {Double} North - Northings in meters
'@param {Double} a     - Ellipsoid semi-major axis in meters
'@param {Double} b     - Ellipsoid semi-minor axis in meters
'@param {Double} e0    - Eastings of false origin in meters
'@param {Double} n0    - Northings of false origin in meters
'@param {Double} f0    - Central meridian scale factor
'@param {Double} PHI0  - Latitude of false origin in decimal degrees
'@returns {Double}     - Convergence in decimal degrees
'@dependencies [InitialLat()]
Private Function E_N_to_C(ByVal East As Double, ByVal North As Double, ByVal a As Double, ByVal b As Double, ByVal E0 As Double, ByVal N0 As Double, ByVal F0 As Double, ByVal phi0 As Double) As Double
  'Convert angle measures to radians
  Dim RadPHI0 As Double: RadPHI0 = phi0 * (PI / 180)
      
  'Compute af0, bf0, e squared (e2), n and Et
  Dim af0 As Double: af0 = a * F0
  Dim bf0 As Double: bf0 = b * F0
  Dim e2 As Double: e2 = ((af0 ^ 2) - (bf0 ^ 2)) / (af0 ^ 2)
  Dim n As Double: n = (af0 - bf0) / (af0 + bf0)
  Dim Et As Double: Et = East - E0

  'Compute initial value for latitude (PHI) in radians
  Dim PHId As Double: PHId = InitialLat(North, N0, af0, RadPHI0, n, bf0)

  'Compute nu, rho and eta2 using value for PHId
  Dim nu As Double: nu = af0 / (Sqr(1 - (e2 * ((Sin(PHId)) ^ 2))))
  Dim rho As Double: rho = (nu * (1 - e2)) / (1 - (e2 * (Sin(PHId)) ^ 2))
  Dim eta2 As Double: eta2 = (nu / rho) - 1

  'Compute Convergence
  Dim XVI As Double: XVI = (Tan(PHId)) / nu
  Dim XVII As Double: XVII = ((Tan(PHId)) / (3 * (nu ^ 3))) * (1 + ((Tan(PHId)) ^ 2) - eta2 - (2 * (eta2 ^ 2)))
  Dim XVIII As Double: XVIII = ((Tan(PHId)) / (15 * (nu ^ 5))) * (2 + (5 * ((Tan(PHId)) ^ 2)) + (3 * ((Tan(PHId)) ^ 4)))

  E_N_to_C = (180 / PI) * ((Et * XVI) - ((Et ^ 3) * XVII) + ((Et ^ 5) * XVIII))
End Function


'Compute local scale factor from latitude and longitude
'@param {Double} PHI   - Latitude in decimal degrees
'@param {Double} LAM   - Longitude in decimal degrees
'@param {Double} LAM0  - Longitude of false origin in decimal degrees
'@param {Double} a     - Ellipsoid semi-major axis in meters
'@param {Double} b     - Ellipsoid semi-minor axis in meters
'@param {Double} f0    - Central meridian scale factor
'@returns {Double}     - the local scaling factor
Private Function Lat_Long_to_LSF(ByVal phi As Double, ByVal LAM As Double, ByVal LAM0 As Double, ByVal a As Double, ByVal b As Double, ByVal F0 As Double) As Double
  'Convert angle measures to radians
  Dim RadPHI As Double: RadPHI = phi * (PI / 180)
  Dim RadLAM As Double: RadLAM = LAM * (PI / 180)
  Dim RadLAM0 As Double: RadLAM0 = LAM0 * (PI / 180)
      
  'Compute af0, bf0 and e squared (e2)
  Dim af0 As Double: af0 = a * F0
  Dim bf0 As Double: bf0 = b * F0
  Dim e2 As Double: e2 = ((af0 ^ 2) - (bf0 ^ 2)) / (af0 ^ 2)

  'Compute nu, rho, eta2 and p
  Dim nu As Double: nu = af0 / (Sqr(1 - (e2 * ((Sin(RadPHI)) ^ 2))))
  Dim rho As Double: rho = (nu * (1 - e2)) / (1 - (e2 * (Sin(RadPHI)) ^ 2))
  Dim eta2 As Double: eta2 = (nu / rho) - 1
  Dim p As Double: p = RadLAM - RadLAM0

  'Compute local scale factor
  Dim XIX As Double: XIX = ((Cos(RadPHI) ^ 2) / 2) * (1 + eta2)
  Dim XX As Double: XX = ((Cos(RadPHI) ^ 4) / 24) * (5 - (4 * ((Tan(RadPHI)) ^ 2)) + (14 * eta2) - (28 * ((Tan(RadPHI * eta2)) ^ 2)))

  Lat_Long_to_LSF = F0 * (1 + ((p ^ 2) * XIX) + ((p ^ 4) * XX))
End Function

'Compute local scale factor from from easting and northing
'@param {Double} East  - Eastings in meters
'@param {Double} North - Northings in meters
'@param {Double} a     - Ellipsoid semi-major axis in meters
'@param {Double} b     - Ellipsoid semi-minor axis in meters
'@param {Double} e0    - Eastings of false origin in meters
'@param {Double} n0    - Northings of false origin in meters
'@param {Double} f0    - Central meridian scale factor
'@param {Double} PHI0  - Latitude of central meridian in decimal degrees
'@returns              - Local scale factor.
'@dependencies [Marc(),InitialLat()]
Private Function E_N_to_LSF(ByVal East As Double, ByVal North As Double, ByVal a As Double, ByVal b As Double, ByVal E0 As Double, ByVal N0 As Double, ByVal F0 As Double, ByVal phi0 As Double) As Double
  'Convert angle measures to radians
  Dim RadPHI0 As Double: RadPHI0 = phi0 * (PI / 180)
      
  'Compute af0, bf0, e squared (e2), n and Et
  Dim af0 As Double: af0 = a * F0
  Dim bf0 As Double: bf0 = b * F0
  Dim e2 As Double: e2 = ((af0 ^ 2) - (bf0 ^ 2)) / (af0 ^ 2)
  Dim n As Double: n = (af0 - bf0) / (af0 + bf0)
  Dim Et As Double: Et = East - E0

  'Compute initial value for latitude (PHI) in radians
  Dim PHId As Double: PHId = InitialLat(North, N0, af0, RadPHI0, n, bf0)

  'Compute nu, rho and eta2 using value for PHId
  Dim nu As Double: nu = af0 / (Sqr(1 - (e2 * ((Sin(PHId)) ^ 2))))
  Dim rho As Double: rho = (nu * (1 - e2)) / (1 - (e2 * (Sin(PHId)) ^ 2))
  Dim eta2 As Double: eta2 = (nu / rho) - 1

  'Compute local scale factor
  Dim XXI As Double: XXI = 1 / (2 * rho * nu)
  Dim XXII As Double: XXII = (1 + (4 * eta2)) / (24 * (rho ^ 2) * (nu ^ 2))

  E_N_to_LSF = F0 * (1 + ((Et ^ 2) * XXI) + ((Et ^ 4) * XXII))
End Function

'Compute (t-T) correction in decimal degrees at point (AtEast, AtNorth) to point (ToEast,ToNorth)
'@param {Double} AtEast  - Eastings in meters, of point where (t-T) is being computed
'@param {Double} AtNorth - Northings in meters, of point where (t-T) is being computed
'@param {Double} ToEast  - Eastings in meters, of point at other end of line to which (t-T) is being computed
'@param {Double} ToNorth - Northings in meters, of point at other end of line to which (t-T) is being computed
'@param {Double} a       - Ellipsoid semi-major axis in meters
'@param {Double} b       - Ellipsoid semi-minor axis in meters
'@param {Double} e0      - Eastings of false origin in meters
'@param {Double} n0      - Northings of false origin in meters
'@param {Double} f0      - Central meridian scale factor
'@param {Double} PHI0    - Latitude of central meridian in decimal degrees
'@returns {Double} correction in decimal degrees
'@dependencies [Marc(),InitialLat()]
Private Function E_N_to_t_minus_T(ByVal AtEast As Double, ByVal AtNorth As Double, ByVal ToEast As Double, ByVal ToNorth As Double, ByVal a As Double, ByVal b As Double, ByVal E0 As Double, ByVal N0 As Double, ByVal F0 As Double, ByVal phi0 As Double) As Double
  'Convert angle measures to radians
  Dim RadPHI0 As Double: RadPHI0 = phi0 * (PI / 180)
      
  'Compute af0, bf0, e squared (e2), n and Nm (Northing of mid point)
  Dim af0 As Double: af0 = a * F0
  Dim bf0 As Double: bf0 = b * F0
  Dim e2 As Double: e2 = ((af0 ^ 2) - (bf0 ^ 2)) / (af0 ^ 2)
  Dim n As Double: n = (af0 - bf0) / (af0 + bf0)
  Dim Nm As Double: Nm = (AtNorth + ToNorth) / 2

  'Compute initial value for latitude (PHI) in radians
  Dim PHId As Double: PHId = InitialLat(Nm, N0, af0, RadPHI0, n, bf0)

  'Compute nu, rho and eta2 using value for PHId
  Dim nu As Double: nu = af0 / (Sqr(1 - (e2 * ((Sin(PHId)) ^ 2))))
  Dim rho As Double: rho = (nu * (1 - e2)) / (1 - (e2 * (Sin(PHId)) ^ 2))

  'Compute (t-T)
  Dim XXIII As Double: XXIII = 1 / (6 * nu * rho)

  E_N_to_t_minus_T = (180 / PI) * ((2 * (AtEast - E0)) + (ToEast - E0)) * (AtNorth - ToNorth) * XXIII
End Function

'Compute true azimuth in decimal degrees at point (AtEast, AtNorth) to point (ToEast,ToNorth)
'@param {Double} AtEast  - Eastings in meters, of point where true azimuth is being computed
'@param {Double} AtNorth - Northings in meters, of point where true azimuth is being computed
'@param {Double} ToEast  - Eastings in meters, of point at other end of line to which true azimuth is being computed
'@param {Double} ToNorth - Northings in meters, of point at other end of line to which true azimuth is being computed
'@param {Double} a       - Ellipsoid semi-major axis in meters
'@param {Double} b       - Ellipsoid semi-minor axis in meters
'@param {Double} e0      - Eastings of false origin in meters
'@param {Double} n0      - Northings of false origin in meters
'@param {Double} f0      - Central meridian scale factor
'@param {Double} PHI0    - Latitude of central meridian in decimal degrees
'@returns {Double}       - True azimuth in decimal degrees
'@dependencies [Marc(),InitialLat(), E_N_to_t_minus_T(), E_N_to_C()]
Private Function TrueAzimuth(ByVal AtEast As Double, ByVal AtNorth As Double, ByVal ToEast As Double, ByVal ToNorth As Double, ByVal a As Double, ByVal b As Double, ByVal E0 As Double, ByVal N0 As Double, ByVal F0 As Double, ByVal phi0 As Double) As Double
  'Compute eastings and northings differences
  Dim Diffe As Double: Diffe = ToEast - AtEast
  Dim Diffn As Double: Diffn = ToNorth - AtNorth

  'Compute grid bearing
  Dim GridBearing As Double
  If Diffe = 0 Then
    If Diffn < 0 Then
      GridBearing = 180
    Else
      GridBearing = 0
    End If
    GoTo EndOfComputeBearing
  End If

  Dim Ratio As Double: Ratio = Diffn / Diffe
  Dim GridAngle As Double: GridAngle = (180 / PI) * Atn(Ratio)

  If Diffe > 0 Then
    GridBearing = 90 - GridAngle
  End If

  If Diffe < 0 Then
    GridBearing = 270 - GridAngle
  End If

EndOfComputeBearing:
  'Compute convergence
  Dim Convergence As Double: Convergence = E_N_to_C(AtEast, AtNorth, a, b, E0, N0, F0, phi0)

  'Compute (t-T) correction
  Dim t_minus_T As Double: t_minus_T = E_N_to_t_minus_T(AtEast, AtNorth, ToEast, ToNorth, a, b, E0, N0, F0, phi0)

  'Compute initial azimuth
  Dim InitAzimuth As Double: InitAzimuth = GridBearing + Convergence - t_minus_T

  'Set TrueAzimuth >=0 and <=360
  If InitAzimuth < 0 Then
    TrueAzimuth = InitAzimuth + 360
  ElseIf InitAzimuth > 360 Then
    TrueAzimuth = InitAzimuth - 360
  Else
    TrueAzimuth = InitAzimuth
  End If
End Function

'**************
'* INTERFACES *
'**************
Private Function gisIPoint_Latitude() As Double:  gisIPoint_Latitude = Latitude: End Function
Private Function gisIPoint_Longitude() As Double: gisIPoint_Longitude = Longitude: End Function
