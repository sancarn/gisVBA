Public rows As Collection

Private Sub Class_Initialize()
    Set rows = New Collection
End Sub



Function From(table As ListObject, Optional base_object As Object) As clsGeo
    'Map "Geography" to object geography
    Dim layer As New clsGeo
    Dim lr As ListRow, row As Collection, sGeo As String
    For Each lr In table.ListRows
        Set row = lrWrap(lr)
        sGeo = row("Geography")
        row.Remove "Geography"
        If base_object Is Nothing Then
            'Test sGeo for type of object
            'row.Add clsGeoPolygon.From(sGeo), "Geography"
        Else
            row.Add base_object.From(sGeo), "Geography"
        End If
        layer.rows.Add row
    Next lr
    Set From = layer
End Function


Private Function lrWrap(lr As ListRow) As Collection
    Dim lo As ListObject
    Set lo = lr.Parent
    Dim vh As Variant: vh = lo.HeaderRowRange.Value 'Header
    Dim vr As Variant: vr = lr.Range.Value          'This row
    Dim retCol As New Collection

    'Append list row and object to collection as __ListRow and __ListObject
    retCol.Add lr, "__ListRow"
    retCol.Add lo, "__ListObject"

    'Loop through each header and append row value with header as key into return collection
    For i = LBound(vh, 2) To UBound(vh, 2)
        retCol.Add vr(1, i), vh(1, i)
    Next

    'Return retCol
    Set lrWrap = retCol
End Function

