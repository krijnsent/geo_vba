Attribute VB_Name = "ModTestGeom"
'Source: https://github.com/krijnsent/geo_vba

Sub RunTests()

'Basic geom test
Dim g As New geom
Debug.Print "g.C_RADIUS_EARTH_KM", g.C_RADIUS_EARTH_KM
Debug.Print "g.Deg2Rad(16)", g.Deg2Rad(16)
Debug.Print "g.ArcSin(0.5)", g.ArcSin(0.5)


'Test simple coordinates
Dim CoordSW As New latLng
CoordSW.lat = 52.067267
CoordSW.lon = 5.1114603
Dim CoordNE As New latLng
CoordNE.lat = 52.074799
CoordNE.lon = 5.1285214

Debug.Print "CoordSW.toString", CoordSW.toString
Debug.Print "CoordNE.toString", CoordNE.toString
Debug.Print "CoordSW.distanceTo(CoordNE)", CoordSW.distanceTo(CoordNE)

Dim CoordTest As New latLng
CoordTest.lat = 52
CoordTest.lon = 5
Debug.Print "CoordTest.toTile 10", CoordTest.toTile(10).toString
Debug.Print "CoordTest.toTile 15", CoordTest.toTile(15).toString


'Test Box/Rectangle/Bounds
Dim TotBox As latLngBounds
Set TotBox = New latLngBounds
TotBox.latLngBounds CoordSW, CoordNE

Debug.Print "TotBox.getCenter.toString", TotBox.getCenter.toString
Debug.Print "TotBox.getNorth", TotBox.getNorth
Debug.Print "TotBox.toString", TotBox.toString

Debug.Print "TotBox.autoZoomLevel(200, 200)", TotBox.autoZoomLevel(200, 200)
Debug.Print "TotBox.autoZoomLevel(500, 500)", TotBox.autoZoomLevel(500, 500)

'One line box
Dim TotBox2 As latLngBounds
Set TotBox2 = New latLngBounds
TotBox2.setBounds Array(Array(51, 5), Array(51.2, 5.1))
Debug.Print "TotBox2.toString", TotBox2.toString


'https://www.openstreetmap.org/#map=16/51.9176/5.8190
'https://tile.openstreetmap.org/16/33830/21672.png
Dim DummyTile As New tile
DummyTile.x = 33830
DummyTile.y = 21672
DummyTile.z = 16

Debug.Print "DummyTile.toString", DummyTile.toString
Set DummyTileLatLng = DummyTile.toLatLng
Debug.Print "DummyTileLatLng.toString", DummyTileLatLng.toString


'Example from https://www.netzwolf.info/osm/tilebrowser.html?tx=34027&ty=21931&tz=16#tile
Dim DummyTile2 As New tile
DummyTile2.x = 34027
DummyTile2.y = 21931
DummyTile2.z = 16

Set DummyTileLatLngBounds = DummyTile2.toLatLngBounds
Debug.Print "DummyTileLatLngBounds.toString", DummyTileLatLngBounds.toString
Debug.Print "DummyTileLatLngBounds.getSurface", DummyTileLatLngBounds.getSurface


Dim DummyTile3 As tile
Set DummyTile3 = New tile
DummyTile3.fromString "16/21931/34027"
Debug.Print "DummyTile3.x", DummyTile3.x
Debug.Print "DummyTile3.y", DummyTile3.y
Debug.Print "DummyTile3.z", DummyTile3.z
Debug.Print "DummyTile3.toString", DummyTile3.toString


'https://stackoverflow.com/questions/28476117/easy-openstreetmap-tile-displaying-for-python
Dim TotBox4 As latLngBounds
Set TotBox4 = New latLngBounds
TotBox4.setBounds Array(Array(38.5, -77.04), Array(38.52, -76.99))
Debug.Print "TotBox4.toString", TotBox4.toString
TileNrsBox = TotBox4.toTiles(13)
For t = LBound(TileNrsBox) To UBound(TileNrsBox)
    Debug.Print "TotBox4 " & t, TileNrsBox(t)
Next t


End Sub

