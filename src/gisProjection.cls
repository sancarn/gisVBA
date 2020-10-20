VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "gisProjection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Enum gisProjectionType
  gisLatLng
  gisBritishNationalGrid
End Enum

Private pName as string

'TODO: Full implementation
Public Function Create(ByVal sName as string) as gisProjection
  set Create = new gisProjection
  Call Create.init(sName)
End Function
Public Sub Init(ByVal sName as string)
  pName = sName
End Sub

Public Property Get Name() as string
  Name = pName
End Property
