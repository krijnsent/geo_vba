Attribute VB_Name = "ModFunctions"
'Source: https://github.com/krijnsent/geo_vba

Private Sub RegisterMyFunction()
    
    Application.MacroOptions _
        Macro:="geo_earth_radius", _
        Description:="Returns the earth radius at the equator.", _
        Category:="Geo_vba formulas", _
        ArgumentDescriptions:=Array( _
            "Optional unit kilometers, miles, nautical miles: km, mi, nm")
    
    Application.MacroOptions _
        Macro:="geo_distance", _
        Description:="Calculates the distance between two coordinates", _
        Category:="Geo_vba formulas", _
        ArgumentDescriptions:=Array( _
            "Latitude (north-south) of coordinate 1, from -90 to +90", _
            "Longitude (east-west) of coordinate 1, from -180 to +180", _
            "Latitude (north-south) of coordinate 2, from -90 to +90", _
            "Longitude (east-west) of coordinate 2, from -180 to +180")
            
    Application.MacroOptions _
        Macro:="geo_surface", _
        Description:="Calculates the surface of the area of two coordinates. Input are two opposite corners of the area, for example the NorthEast and SouthWest corner.", _
        Category:="Geo_vba formulas", _
        ArgumentDescriptions:=Array( _
            "Latitude (north-south) of coordinate 1, from -90 to +90", _
            "Longitude (east-west) of coordinate 1, from -180 to +180", _
            "Latitude (north-south) of coordinate 2, from -90 to +90", _
             "Longitude (east-west) of coordinate 2, from -180 to +180")
           
End Sub

Public Function geo_earth_radius(Optional Unit As String) As Double
Attribute geo_earth_radius.VB_Description = "Returns the earth radius at the equator."
Attribute geo_earth_radius.VB_ProcData.VB_Invoke_Func = " \n21"

'Accepted units: km, mi, nm
Dim g As New geom

If IsMissing(Unit) Then
    geo_earth_radius = g.C_RADIUS_EARTH_KM
Else
    If LCase(Unit) = "km" Or Unit = "" Then
        geo_earth_radius = g.C_RADIUS_EARTH_KM
    ElseIf LCase(Unit) = "mi" Then
        geo_earth_radius = g.C_RADIUS_EARTH_MI
    ElseIf LCase(Unit) = "nm" Then
        geo_earth_radius = g.C_RADIUS_EARTH_NM
    Else
        geo_earth_radius = CVErr(xlErrNum)
    End If
End If

End Function

Public Function geo_distance(lat1 As Double, lon1 As Double, lat2 As Double, lon2 As Double) As Double
Attribute geo_distance.VB_Description = "Calculates the distance between two coordinates"
Attribute geo_distance.VB_ProcData.VB_Invoke_Func = " \n21"

Dim g As New geom
'Check input
If g.IsValidLat(lat1) And g.IsValidLat(lat2) And g.IsValidLon(lon1) And g.IsValidLon(lon2) Then
    'Lats&Lons OK
Else
    geo_distance = CVErr(xlErrNum)
    Exit Function
End If

Dim Coord1 As New latLng
Coord1.lat = lat1
Coord1.lon = lon1
Dim Coord2 As New latLng
Coord2.lat = lat2
Coord2.lon = lon2

geo_distance = Coord1.distanceTo(Coord2)

End Function

Public Function geo_surface(lat1 As Double, lon1 As Double, lat2 As Double, lon2 As Double) As Double
Attribute geo_surface.VB_Description = "Calculates the surface of the area of two coordinates. Input are two opposite corners of the area, for example the NorthEast and SouthWest corner."
Attribute geo_surface.VB_ProcData.VB_Invoke_Func = " \n21"

'Check input
Dim g As New geom
If g.IsValidLat(lat1) And g.IsValidLat(lat2) And g.IsValidLon(lon1) And g.IsValidLon(lon2) Then
    'Lats&Lons OK
Else
    geo_surface = CVErr(xlErrNum)
    Exit Function
End If

'One box
Dim TotBox As latLngBounds
Set TotBox = New latLngBounds
TotBox.setBounds Array(Array(lat1, lon1), Array(lat2, lon2))

geo_surface = TotBox.getSurface

End Function
