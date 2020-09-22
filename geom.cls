VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "geom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Source: https://github.com/krijnsent/geo_vba
'CLASS: geom
'Call by:  Dim g As New geom
'Alternative: application.worksheetfunction.sinh etc.

'Constants:
Public C_RADIUS_EARTH_KM As Double 'Earth radius in kilometers
Public C_RADIUS_EARTH_MI As Double 'Earth radius in miles
Public C_RADIUS_EARTH_NM As Double 'Earth radius in nautical miles
Public C_MI_KM  As Double 'Miles per kilometer
Public C_PI As Double

' Event - triggered when class created
Private Sub Class_Initialize()

    C_RADIUS_EARTH_KM = 6370.97327862
    C_RADIUS_EARTH_MI = 3958.7392618
    C_RADIUS_EARTH_NM = 3437.74677
    
    C_MI_KM = C_RADIUS_EARTH_MI / C_RADIUS_EARTH_KM
    
    C_PI = 4 * Atn(1)
    
End Sub

'Arc functions (ArcSin = inverse Sin)
Public Function ArcSin(x As Double) As Double
      If Abs(x) = 1 Then
        ArcSin = Sgn(x) * Pi / 2
      Else
        ArcSin = Atn(x / Sqr(-x * x + 1))
      End If
End Function
Public Function ArcTan(x As Double) As Double
    ' added for clarity
    ArcTan = Atn(x)
End Function
Public Function ArcCos(x As Double) As Double
    If x = 1 Then
        ArcCos = 0
    ElseIf x = -1 Then
        ArcCos = 4 * Atn(1)
    Else
        ArcCos = Atn(-x / Sqr(-x * x + 1)) + 2 * Atn(1)
    End If
End Function


'Hyperbolic Sin functions
'http://www.devx.com/vb2themax/Tip/19025
Public Function SinH(x As Double) As Double
    Dim temp As Double
    temp = Exp(x)
    SinH = (temp - 1 / temp) / 2
End Function
Function ArcSinH(x As Double) As Double
    ArcSinH = Log(x + Sqr(x * x + 1))
End Function


'Functions: Degrees to Radians and back
Public Function Deg2Rad(value As Double) As Double
    'Convert Degrees to Radians
    Deg2Rad = (value / 180) * C_PI
End Function
Public Function Rad2Deg(value As Double) As Double
    'Convert Radians to Degrees
    Rad2Deg = 180 * (value / C_PI)
End Function


'Functions: MAX and MIN
Public Function Max(value1, value2)
    If value1 > value2 Then
        Max = value1
    Else
        Max = value2
    End If
End Function
Public Function Min(value1, value2)
    If value1 > value2 Then
        Max = value2
    Else
        Max = value1
    End If
End Function

Public Function IsValidLat(value1) As Boolean

IsValidLat = False
If IsNumeric(value1) Then
    If value1 > -90 And value1 < 90 Then
        IsValidLat = True
    End If
End If

End Function

Public Function IsValidLon(value1) As Boolean

IsValidLon = False
If IsNumeric(value1) Then
    If value1 > -180 And value1 <= 180 Then
        IsValidLon = True
    End If
End If

End Function
