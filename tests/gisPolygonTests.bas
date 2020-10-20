

Public Sub zTEST__From()
    'All tests check for some line where x1,y1,x2,y2 = 1,2,3,4
    TEST__ReportFrom "[[1,1],[3,1],[2,2]]", Me.From("[[1,1],[3,1],[2,2]]")
    TEST__ReportFrom """[1,1]"",""[3,1]"",""[2,2]""", Me.From("[1,1]", "[3,1]", "[2,2]")
    TEST__ReportFrom "Arrays of Integers", Me.From(Array(1, 1), Array(3, 1), Array(2, 2))
    TEST__ReportFrom "Arrays of Strings", Me.From(Split("1 1"), Split("3 1"), Split("2 2"))
    TEST__ReportFrom "Collections of Integers", Me.From(TEST__GetCol(Array(TEST__GetCol(Array(1, 1)), TEST__GetCol(Array(3, 1)), TEST__GetCol(Array(2, 2)))))
    TEST__ReportFrom "Collections of Strings", Me.From(TEST__GetCol(Array(TEST__GetCol(Array("1", "1")), TEST__GetCol(Array("3", "1")), TEST__GetCol(Array("2", "2")))))
End Sub
Private Sub TEST__ReportFrom(Test As String, v As Variant)
    On Error Resume Next
        If TypeName(v) = TypeName(Me) Then
            Dim b As Boolean
            b = True
            b = b And v.Points(1).x = 1: b = b And v.Points(1).y = 1
            b = b And v.Points(2).x = 3: b = b And v.Points(2).y = 1
            b = b And v.Points(3).x = 2: b = b And v.Points(3).y = 2
            Debug.Print TypeName(Me) & iff(b, " PASS: ", " FAIL: ") & Test & iff(Not (b), ":: Points incorrect.", ":: Points")
            
            b = True
            b = b And v.Edges(1).Points(1).x = 1: b = b And v.Edges(1).Points(1).y = 1
            b = b And v.Edges(1).Points(2).x = 3: b = b And v.Edges(1).Points(2).y = 1
            Debug.Print TypeName(Me) & iff(b, " PASS: ", " FAIL: ") & Test & iff(Not (b), ":: Line1 incorrect.", ":: Line1")
            
            b = True
            b = b And v.Edges(2).Points(1).x = 3: b = b And v.Edges(2).Points(1).y = 1
            b = b And v.Edges(2).Points(2).x = 2: b = b And v.Edges(2).Points(2).y = 2
            Debug.Print TypeName(Me) & iff(b, " PASS: ", " FAIL: ") & Test & iff(Not (b), ":: Line2 incorrect.", ":: Line2")
            
            b = True
            b = b And v.Edges(3).Points(1).x = 2: b = b And v.Edges(3).Points(1).y = 2
            b = b And v.Edges(3).Points(2).x = 1: b = b And v.Edges(3).Points(2).y = 1
            Debug.Print TypeName(Me) & iff(b, " PASS: ", " FAIL: ") & Test & iff(Not (b), ":: Line3 incorrect.", ":: Line3")
            
        Else
            Debug.Print TypeName(Me) & " FAIL: " & Test & ":: Not a " & TypeName(Me)
        End If
    On Error GoTo 0
End Sub
Private Function TEST__GetCol(ByVal arr As Variant) As Collection
    Dim i As Long, c As New Collection
    For i = LBound(arr) To UBound(arr)
        c.Add arr(i)
    Next i
    Set TEST__GetCol = c
End Function