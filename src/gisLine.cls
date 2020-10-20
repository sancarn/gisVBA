VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "gisLine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Public Points As Collection
Public delta_x As Double
Public delta_y As Double
Public gradient As Double

Private Sub Class_Initialize()
    Set Me.Points = New Collection
End Sub
Private Sub Class_Terminate()
End Sub

Public Sub xPROTECTED_INITIATE_DEFAULTS()
    Me.delta_x = Me.Points(2).x - Me.Points(1).x
    Me.delta_y = Me.Points(2).y - Me.Points(1).y
    
    'Allow for infinity here
    If Me.delta_y = 0 Then
        Me.gradient = getSpecial(1)
    Else
        Me.gradient = Me.delta_x / Me.delta_y
    End If
        
    On Error GoTo 0
End Sub

Public Function Create() As clsGeoLine
    Set Create = Me.From()
End Function
Public Function From(ParamArray Params() As Variant) As clsGeoLine
    Dim ParamsLength As Long
    ParamsLength = UBound(Params) - LBound(Params) + 1
    
    Dim obj As New clsGeoLine
    
    Dim i As Long, v As Variant
    If ParamsLength = 1 Then
        'If type of argument is clsGeoLine, return it.
        If TypeName(Params(0)) = TypeName(Me) Then
            Set From = Params(0)
            Exit Function
        End If
        
        'Parse args provided
        Select Case VarType(Params(0))
            Case Is > vbArray
                v = Params(0)
                obj.Points.Add clsGeoPoint.From(v(LBound(v)))
                obj.Points.Add clsGeoPoint.From(v(UBound(v)))
            Case vbObject 'Assume collection
                If TypeName(Params(0)) = "Collection" Then
                    Set v = Params(0)
                    obj.Points.Add clsGeoPoint.From(v(1))
                    obj.Points.Add clsGeoPoint.From(v(2))
                Else
                    Err.Raise -1, "Line.From", "Incorrect object type parsed to Line.From."
                End If
            Case vbString   'Assume JSON format: [[x1,y1],[x2,y2],[x3,y3],...]
                '#HACK - Parse JSON
                Dim data As Variant
                sData = Replace(Params(0), "[", "")
                data = Split(Mid(sData, 1, Len(sData) - 2), "],")
                
                obj.Points.Add clsGeoPoint.From(Split(data(0), ","))
                obj.Points.Add clsGeoPoint.From(Split(data(1), ","))
        End Select
    ElseIf ParamsLength = 2 Then
        '<<x1,y1>>,<<x2,y2>>
        Set From = Me.From(Params)
        Exit Function
    ElseIf ParamsLength = 4 Then
        'x1,y1,x2,y2
        For i = LBound(Params) To UBound(Params) Step 2
            obj.Points.Add clsGeoPoint.From(Params(i), Params(i + 1))
        Next
    End If
    
    'Set defaults for optimisation
    obj.xPROTECTED_INITIATE_DEFAULTS
    Set From = obj
End Function

Sub zTEST__From()
    'All tests check for some line where x1,y1,x2,y2 = 1,2,3,4
    TEST__ReportFrom "[[1,2],[3,4]]", Me.From("[[1,2],[3,4]]")
    TEST__ReportFrom "Array(1,2),Array(3,4)", Me.From(Array(1, 2), Array(3, 4))
    TEST__ReportFrom "Array(""1"",""2""),Array(""3"",""4"")", Me.From(Split("1 2"), Split("3 4"))
    TEST__ReportFrom "1,2,3,4", Me.From(1, 2, 3, 4)
    TEST__ReportFrom "Collection(1,2),Collection(3,4)", Me.From(TEST__GetCol(Array(TEST__GetCol(Array(1, 2)), TEST__GetCol(Array(3, 4)))))
    TEST__ReportFrom "Collection(""1"",""2""),Collection(""3"",""4"")", Me.From(TEST__GetCol(Array(TEST__GetCol(Array("1", "2")), TEST__GetCol(Array("3", "4")))))
End Sub
Private Sub TEST__ReportFrom(Test As String, v As Variant)
    If TypeName(v) = TypeName(Me) Then
        Dim b As Boolean
        b = True
        b = b And v.Points(1).x = 1: b = b And v.Points(1).y = 2
        b = b And v.Points(2).x = 3: b = b And v.Points(2).y = 4
        Debug.Print TypeName(Me) & iff(b, " PASS: ", " FAIL: ") & Test & iff(Not (b), ":: Values incorrect.", "")
    Else
        Debug.Print TypeName(Me) & " FAIL: " & Test & ":: Not a " & TypeName(Me)
    End If
End Sub
Private Function TEST__GetCol(ByVal arr As Variant) As Collection
    Dim i As Long, c As New Collection
    For i = LBound(arr) To UBound(arr)
        c.Add arr(i)
    Next i
    Set TEST__GetCol = c
End Function


'clsGeoLine.From(0,0,2,2).Intersects(clsGeoLine.From(1,0,0,1)) --> False
Public Function Intersects(ByVal line As clsGeoLine) As Boolean
    Intersects = Not (Me.IntersectsAt(line) Is Nothing)
End Function

Public Function IntersectsAt(ByVal line As clsGeoLine, Optional segmentation As Boolean = True, Optional tolerance As Double = 0) As clsGeoPoint
    '#Let:
    ' Line1 = [[x_1,y_1],[x_2,y_2]]
    ' Line2 = [[x_3,y_3],[x_4,y_4]]
    '#Solving for u:
    'x_1 + t(x_2-x_1) = x_3 + u(x_4-x_3)
    'y_1 + t(y_2-y_1) = y_3 + u(y_4-y_3)
    '#Results in:
    'u = (x_3 (y_2 - y_1) + x_2 (y_1 - y_3) + x_1 (y_3 - y_2))/(x_4 (y_1 - y_2) + x_3 (y_2 - y_1) + (x_1 - x_2) (y_3 - y_4))
    '#Such that:
    'Point of intersection = (x_3 + u(x_4-x_3),y_3 + u(y_4-y_3))
    '#Note: Point of intersection doesn't mean that they intersect on the line segment. However when intersecting on the line
    '#segment 0 <= u <= 1 and 0 <= t <= 1. Hence.
    't = (x_3 - x_1 + u(x_4-x_3))/(x_2-x_1)
    
    Dim x_1, y_1, x_2, y_2, x_3, y_3, x_4, y_4 As Double
    x_1 = Me.Points(1).x
    y_1 = Me.Points(1).y
    x_2 = Me.Points(2).x
    y_2 = Me.Points(2).y
    x_3 = line.Points(1).x
    y_3 = line.Points(1).y
    x_4 = line.Points(2).x
    y_4 = line.Points(2).y
    
    
    'If gradients are equal then lines are parallel thus no intersection
    'TODO: This won't always be correct, although unlikely that it won't be.
    If Me.gradient = line.gradient Then
        Set IntersectsAt = Nothing
        Exit Function
    End If
    
    'Lines which although aren't exactly parallel, are almost parallel (within a tolerance)
    If tolerance <> 0 Then
        If Abs(Me.gradient - line.gradient) < tolerance Then
            Set IntersectsAt = Nothing
            Exit Function
        End If
    End If
    
    'radius of intersect for line_2 per unit line length
    Dim u As Double
    u = (x_3 * (y_2 - y_1) + x_2 * (y_1 - y_3) + x_1 * (y_3 - y_2)) / (x_4 * (y_1 - y_2) + x_3 * (y_2 - y_1) + (x_1 - x_2) * (y_3 - y_4))
    
    'If check wanted to determine the intersection point between 2 line segments, set segmentation = true
    'else the projected vector intersection is returned.
    If segmentation Then
        'radius of intersect for line_1 per unit line length
        Dim t As Double
        
        'CHECK: Not actually sure if this matters, but sometimes t will hit infinity.
        If x_2 = x_1 Then
            If y_2 = y_1 Then
                t = getSpecial(1)
            Else
                t = (y_3 - y_1 + u * (y_4 - y_3)) / (y_2 - y_1)
            End If
        Else
            t = (x_3 - x_1 + u * (x_4 - x_3)) / (x_2 - x_1)
        End If
        
        
        'If Radii between 0 and 1 then they intersect the line segment itself. Else they are a number of unit lengths
        'beyond each line
        If u < 0 Or u > 1 Then
            Set IntersectsAt = Nothing
            Exit Function
        End If
        If t < 0 Or t > 1 Then
            Set IntersectsAt = Nothing
            Exit Function
        End If
    End If
    
    'Intersection point
    Dim px, py As Double
    px = x_3 + u * (x_4 - x_3)
    py = y_3 + u * (y_4 - y_3)
    
    'Return data
    Set IntersectsAt = clsGeoPoint.From(px, py)
End Function

Sub zTest__IntersectsAt()
    On Error Resume Next
        'Intersection should be at 0.5, 0.5
        Dim intersection As clsGeoPoint
        Set intersection = clsGeoLine.From(0, 0, 2, 2).IntersectsAt(clsGeoLine.From(1, 0, 0, 1))
        Debug.Print TypeName(Me) & iff(intersection.x = 0.5 And intersection.y = 0.5, " PASS: IntersectsAt() - Do intersect", " FAIL: IntersectsAt(), incorrect intersection location.")
        
        'Intersection shouldn't exist
        Set intersection = clsGeoLine.From(0, 0, 2, 2).IntersectsAt(clsGeoLine.From(1, 0, 1, 0.5))
        Debug.Print TypeName(Me) & iff(intersection Is Nothing, " PASS: IntersectsAt() - Would but don't intersect", " FAIL: IntersectsAt(), does intersect but shouldn't.")
        
        'Intersection will exist if no segmentation
        Set intersection = clsGeoLine.From(0, 0, 2, 2).IntersectsAt(clsGeoLine.From(1, 0, 1, 0.5), False)
        Debug.Print TypeName(Me) & iff(intersection.x = 1 And intersection.y = 1, " PASS: IntersectsAt() - Vector intersection location correct", " FAIL: IntersectsAt() - Vector intersection location incorrect.")
    On Error GoTo 0
End Sub

Function iff(cond, a, b) As Variant
    If cond Then
        iff = a
    Else
        iff = b
    End If
End Function


Private Function getSpecial(data_in As Double) As Double
    'if data_in =  1 ==> Infinity
    'if data_in =  0 ==> NaN
    'if data_in = -1 ==> -Infinity
    On Error Resume Next
        getSpecial = data_in / 0
    On Error GoTo 0
End Function
