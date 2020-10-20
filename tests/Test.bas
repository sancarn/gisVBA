Sub test_mapping()
    Dim SMSBoundaries As New clsGeo
    Dim HFRRPoints As New clsGeo
    
    
    Set SMSBoundaries = clsGeo.From(Sheet1.ListObjects("SMSBoundaries"), clsGeoPolygon)
    Set HFRRPoints = clsGeo.From(Sheet1.ListObjects("HFRR"), clsGeoPoint)
    
    Dim cSMSBoundary As Collection
    Dim HFRRPoint As Collection
    
    For Each cSMSBoundary In SMSBoundaries.rows
        For Each HFRRPoint In HFRRPoints.rows
            If cSMSBoundary("Geography").Contains(HFRRPoint("Geography")) Then
                Debug.Assert False
                HFRRPoint("__ListRow").Range.Offset(0, 5).Value = cSMSBoundary("SMPRef")
                GoTo SkipRemainingSMPs
            End If
        Next
SkipRemainingSMPs:
    Next
End Sub
