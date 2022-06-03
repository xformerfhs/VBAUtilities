VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BearingHelper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
'+-------------------------------------------------------------------------
'|
'| SPDX-FileCopyrightText: 2022 Frank Schwab
'|
'| SPDX-License-Identifier: MIT
'|
'| Copyright 2022, Frank Schwab
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
'| Class               | BearingHelper
'|---------------------+---------------------------------------------------
'| Description         | Helps with bearing calculations
'|---------------------+---------------------------------------------------
'| Author              | Frank Schwab
'|---------------------+---------------------------------------------------
'| Version             | 2.0.0
'|---------------------+---------------------------------------------------
'| Changes             | 2020-07-20  Created. fhs
'|                     | 2022-06-03  Added bearing calculation,
'|                     |             consistent naming. fhs
'+---------------------+---------------------------------------------------
'

Option Explicit

'
' Private constants
'
Private Const PI As Double = 3.14159265358979
Private Const TWO_PI As Double = PI + PI

Private Const FULL_DEGREES As Double = 360#
Private Const HALF_DEGREES As Double = 180#
Private Const NEGATIVE_HALF_DEGREES As Double = -180#

Private Const DEGREE_TO_RADIANS As Double = PI / HALF_DEGREES
Private Const RADIANS_TO_DEGREE As Double = 1 / DEGREE_TO_RADIANS

'
' Instance variables
'
Private m_TrigoHelper As Trigonometrics

'
' Private methods
'

'
'+--------------------------------------------------------------------------
'| Method           | GetTrigonometrics
'|------------------+-------------------------------------------------------
'| Description      | Get trigonometrics object.
'|------------------+-------------------------------------------------------
'| Parameters       | ./.
'|------------------+-------------------------------------------------------
'| Return values    | Trigonometrics object.
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2022-06-03  Created. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | ./.
'+--------------------------------------------------------------------------
'
Private Function GetTrigonometrics() As Trigonometrics
   If m_TrigoHelper Is Nothing Then _
      m_TrigoHelper = New Trigonometrics

   Set GetTrigonometrics = m_TrigoHelper
End Function

'
' Public methods
'

'
'+--------------------------------------------------------------------------
'| Method           | InBetweenBearingInRadians
'|------------------+-------------------------------------------------------
'| Description      | Get bearing inbetween two bearings in radians.
'|------------------+-------------------------------------------------------
'| Parameters       | startBearing: Start bearing in radians.
'|                  | endBearing: End bearing in radians.
'|                  | fraction: Fraction between start and end bearing.
'|------------------+-------------------------------------------------------
'| Return values    | Bearing at fraction between start and end bearing.
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2022-06-03  Created. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | ./.
'+--------------------------------------------------------------------------
'
Public Function InBetweenBearingInRadians(ByVal startBearing As Double, ByVal endBearing As Double, ByVal fraction As Double) As Double
   InBetweenBearingInRadians = InBetweenBearingInDegrees(startBearing * RADIANS_TO_DEGREE, endBearing * RADIANS_TO_DEGREE, fraction)
End Function

'
'+--------------------------------------------------------------------------
'| Method           | MeanBearingInRadians
'|------------------+-------------------------------------------------------
'| Description      | Get mean bearing between two bearings in radians.
'|------------------+-------------------------------------------------------
'| Parameters       | startBearing: Start bearing in radians.
'|                  | endBearing: End bearing in radians.
'|------------------+-------------------------------------------------------
'| Return values    | Mean bearing between start and end bearing.
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2022-06-03  Created. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | ./.
'+--------------------------------------------------------------------------
'
Public Function MeanBearingInRadians(ByVal startBearing As Double, ByVal endBearing As Double) As Double
   MeanBearingInRadians = InBetweenBearingInRadians(startBearing, endBearing, 0.5)
End Function

'
'+--------------------------------------------------------------------------
'| Method           | InBetweenBearingInDegrees
'|------------------+-------------------------------------------------------
'| Description      | Get bearing inbetween two bearings in degrees.
'|------------------+-------------------------------------------------------
'| Parameters       | startBearing: Start bearing.
'|                  | endBearing: End bearing.
'|                  | fraction: Fraction between start and end bearing.
'|------------------+-------------------------------------------------------
'| Return values    | Bearing at fraction between start and end bearing.
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2022-06-03  Created. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | ./.
'+--------------------------------------------------------------------------
'
Public Function InBetweenBearingInDegrees(ByVal startBearing As Double, ByVal endBearing As Double, ByVal fraction As Double) As Double
   Dim diffBearing As Double

   diffBearing = endBearing - startBearing

   If diffBearing > HALF_DEGREES Then
      diffBearing = diffBearing - FULL_DEGREES
   Else
      If diffBearing < NEGATIVE_HALF_DEGREES Then
         diffBearing = diffBearing + FULL_DEGREES
      End If
   End If

   Dim result As Integer

   result = startBearing + (fraction * diffBearing)

   If result >= FULL_DEGREES Then
      result = result - FULL_DEGREES
   Else
      If result < 0 Then
         result = result + FULL_DEGREES
      End If
   End If

   InBetweenBearing = result
End Function

'
'+--------------------------------------------------------------------------
'| Method           | MeanBearingInDegrees
'|------------------+-------------------------------------------------------
'| Description      | Get mean bearing between two bearings in degrees.
'|------------------+-------------------------------------------------------
'| Parameters       | startBearing: Start bearing in degrees.
'|                  | endBearing: End bearing in degrees.
'|------------------+-------------------------------------------------------
'| Return values    | Mean bearing between start and end bearing.
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2022-06-03  Created. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | ./.
'+--------------------------------------------------------------------------
'
Public Function MeanBearingInDegrees(ByVal startBearing As Double, ByVal endBearing As Double) As Double
   MeanBearingInDegrees = InBetweenBearing(startBearing, endBearing, 0.5)
End Function

'
'+--------------------------------------------------------------------------
'| Method           | BearingInRadians
'|------------------+-------------------------------------------------------
'| Description      | Get the bearing from two points in radians.
'|------------------+-------------------------------------------------------
'| Parameters       | startLatitude: Latitude of start point in radians.
'|                  | startLongitude: Longitude of start point in radians.
'|                  | endLatitude: Latitude of end point in radians.
'|                  | endLongitude: Longitude of end point in radians.
'|------------------+-------------------------------------------------------
'| Return values    | Bearing from start point to end point in radians.
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2022-06-03  Created. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | ./.
'+--------------------------------------------------------------------------
'
Public Function BearingInRadians(ByVal startLatitude As Double, ByVal startLongitude As Double, ByVal endLatitude As Double, ByVal endLongitude As Double) As Double
   Dim diffLongitude As Double
   
   diffLongitude = endLongitude - startLongitude
   
   BearingInRadians = GetTrigonometrics.ArcTan2(Sin(diffLongitude) * Cos(endLatitude), Cos(startLatitude) * Sin(endLatitude) - Sin(startLatitude) * Cos(endLatitude) * Cos(diffLongitude)) + TWO_PI
End Function

'
'+--------------------------------------------------------------------------
'| Method           | BearingInDegrees
'|------------------+-------------------------------------------------------
'| Description      | Get the bearing from two points in degrees.
'|------------------+-------------------------------------------------------
'| Parameters       | startLatitude: Latitude of start point in degrees
'|                  | startLongitude: Longitude of start point in degrees
'|                  | endLatitude: Latitude of end point in degrees
'|                  | endLongitude: Longitude of end point in degrees
'|------------------+-------------------------------------------------------
'| Return values    | Bearing from start point to end point in degrees.
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2022-06-03  Created. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | ./.
'+--------------------------------------------------------------------------
'
Public Function BearingInDegrees(ByVal startLatitude As Double, ByVal startLongitude As Double, ByVal endLatitude As Double, ByVal endLongitude As Double) As Double
   BearingInDegrees = BearingInRadians(startLatitude * DEGREE_TO_RADIANS, startLongitude * DEGREE_TO_RADIANS, endLatitude * DEGREE_TO_RADIANS, endLongitude * DEGREE_TO_RADIANS) * RADIANS_TO_DEGREE
End Function
