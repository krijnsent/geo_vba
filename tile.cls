VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "tile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

' Member variables
Private tx As Long
Private ty As Long
Private tz As Integer

' Properties
Property Get x() As Long
    x = tx
End Property
Property Let x(value As Long)
    tx = value
End Property

Property Get y() As Long
    y = ty
End Property
Property Let y(value As Long)
    ty = value
End Property

Property Get z() As Integer
    z = tz
End Property
Property Let z(value As Integer)
    tz = value
End Property

' Event - triggered when class created
Private Sub Class_Initialize()
    tx = 0
    ty = 0
    tz = 0
End Sub

Public Function toLatLng() As latLng

    Dim g As New geom
    Dim lat_deg As Double
    Dim lon_deg As Double
    Dim lat_rad As Double
    
    n = 2 ^ tz
    lon_deg = tx / n * 360 - 180
    lat_rad = Atn(g.SinH(g.C_PI * (1 - 2 * ty / n)))
    lat_deg = g.Rad2Deg(lat_rad)
    
    Dim Coord As New latLng
    Coord.lat = lat_deg
    Coord.lon = lon_deg

    Set toLatLng = Coord

End Function

Public Function toLatLngBounds() As latLngBounds

    Dim g As New geom
    Dim lat_deg As Double, lat_deg2 As Double
    Dim lon_deg As Double, lon_deg2 As Double
    Dim lat_rad As Double, lat_rad2 As Double
    
    Dim CoordBounds As New latLngBounds
    
    n = 2 ^ tz
    lon_deg = tx / n * 360 - 180
    lat_rad = Atn(g.SinH(g.C_PI * (1 - 2 * ty / n)))
    lat_deg = g.Rad2Deg(lat_rad)
    
    lon_deg2 = (tx + 1) / n * 360 - 180
    lat_rad2 = Atn(g.SinH(g.C_PI * (1 - 2 * (ty + 1) / n)))
    lat_deg2 = g.Rad2Deg(lat_rad2)
        
    CoordBounds.setBounds Array(Array(lat_deg, lon_deg), Array(lat_deg2, lon_deg2))

    Set toLatLngBounds = CoordBounds

End Function


Public Function toString() As String

    'map tiles format
    toString = "" & tz & "/" & tx & "/" & ty

End Function

Public Sub fromString(StrIn As String)
    'Assuming in: z/x/y
        
    Dim t As New tile
    CoordArr = Split(StrIn, "/")
    If IsArray(CoordArr) Then
        If UBound(CoordArr) = 2 Then
            tx = CoordArr(1)
            ty = CoordArr(2)
            tz = CoordArr(0)
        End If
    End If
    
End Sub
