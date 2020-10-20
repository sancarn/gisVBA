VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "gisLayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Enum gisFileType
  gisGeoJSON
End Enum
Public Enum gisOperation
  layerObjectWithin
  layerObjectCentroidWithin
  layerObjectContains
  layerObjectContainsCentroid
  layerObjectIntersects
End Enum


'TODO: Build a QuadTree for increased performance

Public Function CreateFromFile(ByVal sFilePath as string, optional ByVal proj as gisProjectionType = gisLatLng, optional ByVal filetype as gisFileType = gisGeoJSON) as gisLayer

End Function

'@param {ByVal Collection<IFeature>|Array<IFeature>}
Public Function CreateFromFeatures(ByVal col as variant) as gisLayer

End Function

Public Sub AddFeature(ByVal feature as gisIFeature)

End Sub
Public Sub RemoveFeature()

End Sub
Public Property Get Feature(ByVal i as long) as gisIFeature

End Property
Public Property Set Feature(ByVal i as long, ByVal feature as gisIFeature) 

End Property

Public Sub ToShapes(ByVal ws as worksheet)

End Sub

'Find based on an ICallable
Public Function FindAll(ByVal callable as stdICallable) as Collection

End Function
Public Function FindFirst(ByVal callable as stdICallable) as gisIFeature

End Function
Public Function FindWhere(ByVal operation as gisOperation, ByVal features as variant)

End Function

Public Function Update()

End Function

Public Function SaveAs(ByVal sSaveToPath as string, Optional ByVal saveAsType as gisFileType = gisGeoJSON) as boolean

End Function

'