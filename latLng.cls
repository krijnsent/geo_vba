VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "latLng"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Source: https://github.com/krijnsent/geo_vba
' CLASS MODULE CODE for latLng
' following logic from https://leafletjs.com/reference-1.7.1.html#latlngbounds
'x = lon - East/West
'y = lat - North/South

' Member variable
Private clat As Double
Private clon As Double

' Properties
Property Get lat() As Double
    lat = clat
End Property
Property Let lat(value As Double)
    clat = value
End Property

Property Get lon() As Double
    lon = clon
End Property

Property Let lon(value As Double)
    clon = value
End Property

' Event - triggered when class created
Private Sub Class_Initialize()
    clat = 0
    clon = 0
End Sub

Public Function toString() As String
    'toString()  String Returns a string representation of the point (for debugging purposes).
    'e.g.  LatLng(51.504789, 5.046692)
    toString = "LatLng(" & Trim(Str(clat)) & "," & Str(clon) & ")"
End Function

Public Function distanceTo(latLngIn As latLng) As Double
    
    'distanceTo(<LatLng> otherLatLng) Number Returns the distance (in meters) to the given LatLng calculated using the Spherical Law of Cosines.
    'http://www.cpearson.com/excel/LatLong.aspx
    Dim Delta As Double
    Dim g As New geom
    
    ' convert to radians
    Dim lat1 As Double
    Dim lat2 As Double
    Dim Long1 As Double
    Dim Long2 As Double
    lat1 = g.Deg2Rad(latLngIn.lat)
    lat2 = g.Deg2Rad(clat)
    Long1 = g.Deg2Rad(latLngIn.lon)
    Long2 = g.Deg2Rad(clon)
    
    Delta = ((2 * g.ArcSin(Sqr((Sin((lat1 - lat2) / 2) ^ 2) + Cos(lat1) * Cos(lat2) * (Sin((Long1 - Long2) / 2) ^ 2)))))
    distanceTo = 1000 * g.C_RADIUS_EARTH_KM * Delta

End Function

Public Function bearingsTo(latLngIn As latLng) As Double

    'Direction in degrees from one coordinate to another, TODO


End Function

Public Function toTile(zoom) As tile
    
    Dim g As New geom
    Dim t As New tile
    
    lat_rad = g.Deg2Rad(lat)
    n = 2 ^ zoom
    xtile = Int((lon + 180) / 360 * n)
    ytile = Int((1 - g.ArcSinH(Tan(lat_rad)) / g.C_PI) / 2 * n)
    
    t.x = xtile
    t.y = ytile
    t.z = zoom
    
    Set toTile = t

End Function
