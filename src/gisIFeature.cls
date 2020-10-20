VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "gisIFeature"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = false
Attribute VB_Exposed = False

Public Property Get Data(ByVal sField as string) as variant

End Property
Public Property Let Data(ByVal sField as string, ByVal v as variant)

End Property

Public Property Get Centroid() as gisPoint

End Property
Public Property Get CentroidX() as Double

End Property
Public Property Get CentroidY() as Double

End Property

'@returns {Collection<gisPoint>}
Public Property Get Points() as Collection

End Property
