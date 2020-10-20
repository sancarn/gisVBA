VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "gisPoint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Public x As Double
Public y As Double

'Private class methods
Private Sub Class_Initialize()
End Sub
Private Sub Class_Terminate()
End Sub



'Public constructor methods
Public Function Create() As clsGeoPoint
    
End Function
Public Function From(ParamArray Params() As Variant) As clsGeoPoint
    Dim ParamsLength As Long
    ParamsLength = UBound(Params) - LBound(Params) + 1
    
    Dim obj As New clsGeoPoint
    If ParamsLength = 1 Then
        'If clsGeoPoint given, return it as required
        If TypeName(Params(0)) = TypeName(Me) Then
            Set From = Params(0)
            Exit Function
        End If
        
        Select Case VarType(Params(0))
        Case vbString
            'Assume JSON
            If Params(0) Like "[[]*[]]" Then
                Set From = clsGeoPoint.From(Split(Replace(Replace(Params(0), "[", ""), "]", ""), ","))
                Exit Function
            Else
                Err.Raise -1, "Point.From", "Params(1) is string, but is not in JSON format."
                Exit Function
            End If
        Case Is > vbArray
            'Assume array of 2 values [x,y]
            obj.x = CDbl(Params(0)(LBound(Params(0))))
            obj.y = CDbl(Params(0)(UBound(Params(0))))
        Case vbObject
            'Assume collection of 2 values [x,y]
            If TypeName(Params(0)) = "Collection" Then
                obj.x = CDbl(Params(0)(1))
                obj.y = CDbl(Params(0)(2))
            Else
                Err.Raise -1, "Point.From", "Invalid object passed to Point.From."
                Exit Function
            End If
        Case Else
            Err.Raise -1, "Point.From", "Unhandled data type."
            Exit Function
        End Select
    ElseIf ParamsLength = 2 Then
        Set From = Me.From(Params)
        Exit Function
    Else
        Err.Raise -1, "Point.From", "Unhandled number of parameters exceeds 2."
        Exit Function
    End If
    
    
    Set From = obj
End Function




