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
'| Class               | BearingHelper
'|---------------------+---------------------------------------------------
'| Description         | Helps with bearing calculations
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
' Public methods
'
Public Function CalculateInBetweenBearing(ByVal bearingFrom As Integer, ByVal bearingTo As Integer, ByVal fraction As Double) As Integer
   Dim diffBearing As Integer

   diffBearing = bearingTo - bearingFrom

   If diffBearing > 180 Then
      diffBearing = diffBearing - 360
   Else
      If diffBearing < -180 Then
         diffBearing = diffBearing + 360
      End If
   End If

   Dim result As Integer

   result = CInt(Round(bearingFrom + (fraction * diffBearing), 0))

   If result >= 360 Then
      result = result - 360
   Else
      If result < 0 Then
         result = result + 360
      End If
   End If

   CalculateInBetweenBearing = result
End Function

Public Function CalculateMeanBearing(ByVal bearingFrom As Integer, ByVal bearingTo As Integer) As Integer
   CalculateMeanBearing = calculateInBetweenBearing(bearingFrom, bearingTo, 0.5)
End Function
