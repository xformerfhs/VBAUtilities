VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SphereDistanceCalculator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'
' Private constants
'
Private Const PI As Double = 3.14159265358979
Private Const TWO_PI As Double = PI + PI

Private Const EARTH_RADIUS_IN_METERS As Double = 6378388

'
' Instance variables
'
Private m_Trigo As New Trigonometrics

'
' Public methods
'
Public Function getDistanceForRadiants(ByVal fromLatitudeRad As Double, _
                                       ByVal fromLongitudeRad As Double, _
                                       ByVal toLatitudeRad As Double, _
                                       ByVal toLongitudeRad As Double, _
                                       ByVal sphereRadius As Double) As Double
   getDistanceForRadiants = sphereRadius * m_Trigo.ArcCos(Sin(fromLatitudeRad) * Sin(toLatitudeRad) + _
                                                          Cos(fromLatitudeRad) * Cos(toLatitudeRad) * Cos(fromLongitudeRad - toLongitudeRad))
End Function

Public Function getDistanceForDegrees(ByVal fromLatitudeDeg As Double, _
                                      ByVal fromLongitudeDeg As Double, _
                                      ByVal toLatitudeDeg As Double, _
                                      ByVal toLongitudeDeg As Double, _
                                      ByVal sphereRadius As Double) As Double
   getDistanceForDegrees = Me.getDistanceForRadiants(m_Trigo.DegreeToRadiant(fromLatitudeDeg), _
                                                     m_Trigo.DegreeToRadiant(fromLongitudeDeg), _
                                                     m_Trigo.DegreeToRadiant(toLatitudeDeg), _
                                                     m_Trigo.DegreeToRadiant(toLongitudeDeg), _
                                                     sphereRadius)
End Function

Public Function getEarthDistanceForRadiants(ByVal fromLatitudeRad As Double, _
                                            ByVal fromLongitudeRad As Double, _
                                            ByVal toLatitudeRad As Double, _
                                            ByVal toLongitudeRad As Double) As Double
   getEarthDistanceForRadiants = Me.getDistanceForRadiants(fromLatitudeRad, fromLongitudeRad, toLatitudeRad, toLongitudeRad, EARTH_RADIUS_IN_METERS)
End Function

Public Function getEarthDistanceForDegrees(ByVal fromLatitudeDeg As Double, _
                                           ByVal fromLongitudeDeg As Double, _
                                           ByVal toLatitudeDeg As Double, _
                                           ByVal toLongitudeDeg As Double) As Double
   getEarthDistanceForDegrees = Me.getDistanceForDegrees(fromLatitudeDeg, fromLongitudeDeg, toLatitudeDeg, toLongitudeDeg, EARTH_RADIUS_IN_METERS)
End Function

Public Function getBearingForRadiants(ByVal fromLatitudeRad As Double, _
                                      ByVal fromLongitudeRad As Double, _
                                      ByVal toLatitudeRad As Double, _
                                      ByVal toLongitudeRad As Double) As Double
   Dim diffLongitude As Double
   
   diffLongitude = toLongitudeRad - fromLongitudeRad
   
   Dim result As Double
   
   result = m_Trigo.ArcTan2(Cos(fromLatitudeRad) * Sin(toLatitudeRad) - Sin(fromLatitudeRad) * Cos(toLatitudeRad) * Cos(diffLongitude), _
                            Sin(diffLongitude) * Cos(toLatitudeRad))
   If result < 0# Then _
      result = TWO_PI + result

   getBearingForRadiants = result
End Function

Public Function getBearingForDegrees(ByVal fromLatitudeDeg As Double, _
                                     ByVal fromLongitudeDeg As Double, _
                                     ByVal toLatitudeDeg As Double, _
                                     ByVal toLongitudeDeg As Double) As Integer
   getBearingForDegrees = CInt(Round(m_Trigo.RadiantToDegree(Me.getBearingForRadiants(m_Trigo.DegreeToRadiant(fromLatitudeDeg), _
                                                                                      m_Trigo.DegreeToRadiant(fromLongitudeDeg), _
                                                                                      m_Trigo.DegreeToRadiant(toLatitudeDeg), _
                                                                                      m_Trigo.DegreeToRadiant(toLongitudeDeg))), _
                                     0))
End Function
