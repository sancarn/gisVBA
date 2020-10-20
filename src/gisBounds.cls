VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "gisBounds"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Public min_x As Double
Public min_y As Double
Public max_x As Double
Public max_y As Double

'Class private initialisation


Private Sub Class_Initialize()
    Call init
End Sub
Private Sub init()
    Dim Infinity, NegInfinity As Double
    Infinity = getSpecial(1)
    NegInfinity = getSpecial(-1)
    
    Me.max_x = NegInfinity
    Me.max_y = NegInfinity
    Me.min_x = Infinity
    Me.min_y = Infinity
End Sub
Private Function getSpecial(data_in As Double) As Double
    'if data_in =  1 ==> Infinity
    'if data_in =  0 ==> NaN
    'if data_in = -1 ==> -Infinity
    On Error Resume Next
        getSpecial = data_in / 0
    On Error GoTo 0
End Function

'Public Methods
Public Function Contains(point As clsGeoPoint) As Boolean
    Contains = point.x >= Me.min_x And point.x <= Me.max_x And point.y >= Me.min_y And point.y <= Me.max_y
End Function
Public Sub Reinitialize()
    Call init
End Sub
Public Sub RefineFrom(point As clsGeoPoint)
    Me.max_x = Max(Me.max_x, point.x)
    Me.max_y = Max(Me.max_y, point.y)
    Me.min_x = Min(Me.min_x, point.x)
    Me.min_y = Min(Me.min_y, point.y)
End Sub
Private Function Max(a As Double, b As Double) As Double
    Max = b
    If a > b Then Max = a
End Function
Private Function Min(a As Double, b As Double) As Double
    Min = b
    If a < b Then Min = a
End Function



