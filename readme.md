# gisVBA

This library is intended to add a GIS engine over functionality within Excel, to help users and developers alike perform GIS style analysis from within Excel.

Long term we'd love to add rendering into this set of libraries, however this will likely not be available until we either implement a VBA http server and use Leaflet as a renderer, or have some sort of canvas control - which is doable but out of scope. This will likely be a task for stdVBA not here.

See the [road map](https://github.com/sancarn/gisVBA/projects/1) for progress details.

This set of libraries heavily relies on functionality supplied by [stdVBA](https://github.com/sancarn/gisVBA).

## Examples

Note this library is heavily work in progress and this example is unlikely to work and will likely change.

### Example 1 - Finding a which region a point is in:

```vb
'Open region layer:
Dim layer as gisLayer
set layer = gisLayer.CreateFromFile("C:\regions.json",gisBritishNationalGrid)

Dim results as Collection
set results = layer.findWhere(layerObjectContains, gisPoint.CreateFromXY(123456,123456))
Debug.Print results(1).data("id")

'or zones:
Dim res as variant
for each res in results
  Debug.Print res.data("id") 
next
```

### Example 2

```vb
'Open a layer from GeoJSON
Dim layer as gisLayer
set layer = gisLayer.CreateFromFile("C:\MyFile.json", gisBritishNationalGrid)

'Add a feature to a layer
layer.addFeature(gisPolygon.CreateFrom1DArray(array(123456,123456,123460,123460,123460,123456)))

'Find all features that match a condition:
Dim bigAreas as collection
set bigAreas = gisLayer.FindAll(stdLambda.Create("$1.area > 4000000"))

'Find areas not at risk
With gisLayer.CreateFromFeatures(bigAreas)
    Dim atRiskZones as gisLayer
    set atRiskZones = gisLayer.CreateFromGeoJSON("C:\RiskZones.json", gisBritishNationalGrid)

    'Save at risk zones to disk:
    with gisLayer.CreateFromFeatures(bigAreas.FindWhere(gisOperations.isCentroidWithin, atRiskZones))
        Call .saveAs("C:\MyBigRiskyZones.json")
    end with
End With
```