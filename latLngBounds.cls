VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "latLngBounds"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' CLASS MODULE CODE for latLngBounds - basically a rectangle on a map
' following logic from https://leafletjs.com/reference-1.7.1.html#latlngbounds
' https://excelmacromastery.com/vba-class-modules/

' Member variable
Private clatmax As Double
Private clonmax As Double
Private clatmin As Double
Private clonmin As Double

' Properties
Public Sub latLngBounds(coord1 As latLng, coord2 As latLng)
    
    If coord1.lat > coord2.lat Then
        clatmax = coord1.lat
        clatmin = coord2.lat
    Else
        clatmax = coord2.lat
        clatmin = coord1.lat
    End If

    If coord1.lon > coord2.lon Then
        clonmax = coord1.lon
        clonmin = coord2.lon
    Else
        clonmax = coord2.lon
        clonmin = coord1.lon
    End If

End Sub

Public Sub setBounds(ArrIn As Variant)

'Expecting one array containing 2 arrays with both 2 numbers, check for that first
If IsArray(ArrIn) Then
    If UBound(ArrIn) = 1 Then
       coord1 = ArrIn(0)
       coord2 = ArrIn(1)
       If IsArray(coord1) And IsArray(coord2) Then
            If UBound(coord1) = 1 And UBound(coord2) = 1 Then
                
                If coord1(0) > coord2(0) Then
                    clatmax = coord1(0)
                    clatmin = coord2(0)
                Else
                    clatmax = coord2(0)
                    clatmin = coord1(0)
                End If
            
                If coord1(1) > coord2(1) Then
                    clonmax = coord1(1)
                    clonmin = coord2(1)
                Else
                    clonmax = coord2(1)
                    clonmin = coord1(1)
                End If
                
            End If
       End If
    End If
End If

End Sub

' Event - triggered when class created
Private Sub Class_Initialize()
    clatmin = 0
    clatmax = 0
    clonmin = 0
    clonmax = 0
End Sub

Public Function getCenter() As latLng
    'getCenter() LatLng
    'Returns the center point of the bounds.
    Dim CoordCenter As New latLng
    CoordCenter.lat = (clatmax + clatmin) / 2
    CoordCenter.lon = (clonmax + clonmin) / 2
    Set getCenter = CoordCenter
End Function

Public Function getWest() As Double
    getWest = clonmin
End Function
Public Function getEast() As Double
    getEast = clonmax
End Function
Public Function getNorth() As Double
    getNorth = clatmax
End Function
Public Function getSouth() As Double
    getSouth = clatmin
End Function

Public Function toString() As String

    'geoJSON format
    toString = "LatLngBounds [[" & Trim(Str(clonmin)) & "," & Str(clatmin) & "]"
    toString = toString & ",[" & Trim(Str(clonmax)) & "," & Str(clatmin) & "]"
    toString = toString & ",[" & Trim(Str(clonmax)) & "," & Str(clatmax) & "]"
    toString = toString & ",[" & Trim(Str(clonmin)) & "," & Str(clatmax) & "]"
    toString = toString & ",[" & Trim(Str(clonmin)) & "," & Str(clatmin) & "]]"

End Function

Public Function getSurface() As Double
    
    Dim g As New geom
        
    clonminr = g.Deg2Rad(clonmin)
    clonmaxr = g.Deg2Rad(clonmax)
    
    clatminr = g.Deg2Rad(clatmin)
    clatmaxr = g.Deg2Rad(clatmax)
    
    getSurface = Abs(clonmaxr - clonminr) * Abs(Sin(clatminr) - Sin(clatmaxr)) * (g.C_RADIUS_EARTH_KM * 1000) ^ 2

End Function

Public Function autoZoomLevel(pxWidth As Integer, pxHeight As Integer) As Double

    'calculate Zoom level to fit the box
    'https://gis.stackexchange.com/questions/19632/how-to-calculate-the-optimal-zoom-level-to-display-two-or-more-points-on-a-map
    
    Dim g As New geom
    Dim ry1 As Double, ry2 As Double, ryc As Double, centery As Double
    
    paddingFactor = 1.1
    
    ry1 = Log((Sin(g.Deg2Rad(clatmin)) + 1) / Cos(g.Deg2Rad(clatmin)))
    ry2 = Log((Sin(g.Deg2Rad(clatmax)) + 1) / Cos(g.Deg2Rad(clatmax)))
    ryc = (ry1 + ry2) / 2
    centery = g.Rad2Deg(Atn(g.SinH(ryc)))
    
    resolutionHorizontal = (clonmax - clonmin) / pxWidth
    vy0 = Log(Tan(g.C_PI * (0.25 + centery / 360)))
    vy1 = Log(Tan(g.C_PI * (0.25 + clatmax / 360)))
    
    viewHeightHalf = pxHeight / 2
    zoomFactorPowered = viewHeightHalf / (40.7436654315252 * (vy1 - vy0))
    resolutionVertical = 360# / (zoomFactorPowered * 256)

    resolution = g.Max(resolutionHorizontal, resolutionVertical) * paddingFactor
    
    autoZoomLevel = Log(360 / (resolution * 256)) / Log(2)

End Function

Function toTiles(zoomLevel As Integer) As String()

'Get an array of tile numbers
Dim tileStrTemp() As String

Dim CoordSW As New latLng
CoordSW.lat = clatmin
CoordSW.lon = clonmin
Set TileSW = CoordSW.toTile(zoomLevel)

Dim CoordNE As New latLng
CoordNE.lat = clatmax
CoordNE.lon = clonmax
Set TileNE = CoordNE.toTile(zoomLevel)

n = 0
For xtile = TileSW.x To TileNE.x + 1
    For ytile = TileNE.y To TileSW.y + 1
        ReDim Preserve tileStrTemp(n)
        tileStrTemp(n) = zoomLevel & "/" & xtile & "/" & ytile
        n = n + 1
    Next ytile
Next xtile

toTiles = tileStrTemp

'https://stackoverflow.com/questions/28476117/easy-openstreetmap-tile-displaying-for-python
'def getImageCluster(lat_deg, lon_deg, delta_lat, delta_long, zoom):
'    smurl = r"http://a.tile.openstreetmap.org/{0}/{1}/{2}.png"
'    xmin , ymax = deg2num(lat_deg, lon_deg, zoom)
'    xmax , ymin = deg2num(lat_deg + delta_lat, lon_deg + delta_long, zoom)
'
'    Cluster = Image.new('RGB',((xmax-xmin+1)*256-1,(ymax-ymin+1)*256-1) )
'    for xtile in range(xmin, xmax+1):
'        for ytile in range(ymin,  ymax+1):
'try:
'                imgurl = smurl.Format(zoom, xtile, ytile)
'                Print ("Opening: " + imgurl)
'                imgstr = urllib2.urlopen(imgurl).Read()
'                tile = Image.Open(StringIO.StringIO(imgstr))
'                Cluster.paste(tile, box=((xtile-xmin)*256 ,  (ytile-ymin)*255))
'except:
'                Print ("Couldn't download image")
'                tile = None
'
'    return Cluster




''At zoom level 9, your world has 2^9 by 2^9 tiles (that's 512x512 tiles).
'def deg2num(lat_deg, lon_deg, zoom):
'  lat_rad = Math.Radians(lat_deg)
'  n = 2.0 ** zoom
'  xtile = Int((lon_deg + 180#) / 360# * n)
'  ytile = Int((1# - Math.Asinh(Math.Tan(lat_rad)) / Math.Pi) / 2# * n)
'  return (xtile, ytile)



End Function

'Factory Description
'L.latLngBounds(<LatLng> corner1, <LatLng> corner2)  Creates a LatLngBounds object by defining two diagonally opposite corners of the rectangle.
'L.latLngBounds(<LatLng[]> latlngs)  Creates a LatLngBounds object defined by the geographical points it contains. Very useful for zooming the map to fit a particular set of locations with fitBounds.
'Methods
'Method  Returns Description
'extend(<LatLng> latlng) this
'Extend the bounds to contain the given point
'
'extend(<LatLngBounds> otherBounds)  this
'Extend the bounds to contain the given bounds
'
'pad(<Number> bufferRatio)   LatLngBounds
'Returns bounds created by extending or retracting the current bounds by a given ratio in each direction. For example, a ratio of 0.5 extends the bounds by 50% in each direction. Negative values will retract the bounds.
'
'getCenter() LatLng
'Returns the center point of the bounds.
'
'getSouthWest()  LatLng
'Returns the south-west point of the bounds.
'
'getNorthEast()  LatLng
'Returns the north-east point of the bounds.
'
'getNorthWest()  LatLng
'Returns the north-west point of the bounds.
'
'getSouthEast()  LatLng
'Returns the south-east point of the bounds.
'
'getWest()   Number
'Returns the west longitude of the bounds
'
'getSouth()  Number
'Returns the south latitude of the bounds
'
'getEast()   Number
'Returns the east longitude of the bounds
'
'getNorth()  Number
'Returns the north latitude of the bounds
'
'contains(<LatLngBounds> otherBounds)    Boolean
'Returns true if the rectangle contains the given one.
'
'contains(<LatLng> latlng)   Boolean
'Returns true if the rectangle contains the given point.
'
'intersects(<LatLngBounds> otherBounds)  Boolean
'Returns true if the rectangle intersects the given bounds. Two bounds intersect if they have at least one point in common.
'
'overlaps(<Bounds> otherBounds)  Boolean
'Returns true if the rectangle overlaps the given bounds. Two bounds overlap if their intersection is an area.
'
'toBBoxString()  String
'Returns a string with bounding box coordinates in a 'southwest_lng,southwest_lat,northeast_lng,northeast_lat' format. Useful for sending requests to web services that return geo data.
'
'equals(<LatLngBounds> otherBounds, <Number> maxMargin?) Boolean
'Returns true if the rectangle is equivalent (within a small margin of error) to the given bounds. The margin of error can be overridden by setting maxMargin to a small number.
'
'isValid()   Boolean
'Returns true if the bounds are properly initialized.
