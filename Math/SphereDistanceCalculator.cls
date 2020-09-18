VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SphereDistanceCalculator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
'+-------------------------------------------------------------------------
'|
'| SPDX-FileCopyrightText: 2020 Frank Schwab
'|
'| SPDX-License-Identifier: MIT
'|
'| Copyright 2020, Frank Schwab
'|
'| Permission is hereby granted, free of charge, to any person obtaining a
'| copy of this software and associated documentation files (the "Software"),
'| to deal in the Software without restriction, including without limitation
'| the rights to use, copy, modify, merge, publish, distribute, sublicense,
'| and/or sell copies of the Software, and to permit persons to whom the
'| Software is furnished to do so, subject to the following conditions:
'|
'| The above copyright notice and this permission notice shall be included
'| in all copies or substantial portions of the Software.
'|
'| THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS
'| OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
'| FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL
'| THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
'| LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
'| OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS
'| IN THE SOFTWARE.
'|
'|-------------------------------------------------------------------------
'| Class               | SphereDistanceCalculator
'|---------------------+---------------------------------------------------
'| Description         | Calculate distances between points on a sphere
'|---------------------+---------------------------------------------------
'| Author              | Frank Schwab
'|---------------------+---------------------------------------------------
'| Version             | 1.0.0
'|---------------------+---------------------------------------------------
'| Changes             | 2020-07-20  Created. fhs
'|---------------------+---------------------------------------------------
'

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
Public Function GetDistanceForRadiants(ByVal fromLatitudeRad As Double, _
                                       ByVal fromLongitudeRad As Double, _
                                       ByVal toLatitudeRad As Double, _
                                       ByVal toLongitudeRad As Double, _
                                       ByVal sphereRadius As Double) As Double
   GetDistanceForRadiants = sphereRadius * m_Trigo.ArcCos(Sin(fromLatitudeRad) * Sin(toLatitudeRad) + _
                                                          Cos(fromLatitudeRad) * Cos(toLatitudeRad) * Cos(fromLongitudeRad - toLongitudeRad))
End Function

Public Function GetDistanceForDegrees(ByVal fromLatitudeDeg As Double, _
                                      ByVal fromLongitudeDeg As Double, _
                                      ByVal toLatitudeDeg As Double, _
                                      ByVal toLongitudeDeg As Double, _
                                      ByVal sphereRadius As Double) As Double
   GetDistanceForDegrees = Me.GetDistanceForRadiants(m_Trigo.DegreeToRadiant(fromLatitudeDeg), _
                                                     m_Trigo.DegreeToRadiant(fromLongitudeDeg), _
                                                     m_Trigo.DegreeToRadiant(toLatitudeDeg), _
                                                     m_Trigo.DegreeToRadiant(toLongitudeDeg), _
                                                     sphereRadius)
End Function

Public Function GetEarthDistanceForRadiants(ByVal fromLatitudeRad As Double, _
                                            ByVal fromLongitudeRad As Double, _
                                            ByVal toLatitudeRad As Double, _
                                            ByVal toLongitudeRad As Double) As Double
   GetEarthDistanceForRadiants = Me.GetDistanceForRadiants(fromLatitudeRad, fromLongitudeRad, toLatitudeRad, toLongitudeRad, EARTH_RADIUS_IN_METERS)
End Function

Public Function GetEarthDistanceForDegrees(ByVal fromLatitudeDeg As Double, _
                                           ByVal fromLongitudeDeg As Double, _
                                           ByVal toLatitudeDeg As Double, _
                                           ByVal toLongitudeDeg As Double) As Double
   GetEarthDistanceForDegrees = Me.GetDistanceForDegrees(fromLatitudeDeg, fromLongitudeDeg, toLatitudeDeg, toLongitudeDeg, EARTH_RADIUS_IN_METERS)
End Function

Public Function GetBearingForRadiants(ByVal fromLatitudeRad As Double, _
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

   GetBearingForRadiants = result
End Function

Public Function GetBearingForDegrees(ByVal fromLatitudeDeg As Double, _
                                     ByVal fromLongitudeDeg As Double, _
                                     ByVal toLatitudeDeg As Double, _
                                     ByVal toLongitudeDeg As Double) As Integer
   GetBearingForDegrees = CInt(Round(m_Trigo.RadiantToDegree(Me.GetBearingForRadiants(m_Trigo.DegreeToRadiant(fromLatitudeDeg), _
                                                                                      m_Trigo.DegreeToRadiant(fromLongitudeDeg), _
                                                                                      m_Trigo.DegreeToRadiant(toLatitudeDeg), _
                                                                                      m_Trigo.DegreeToRadiant(toLongitudeDeg))), _
                                     0))
End Function
