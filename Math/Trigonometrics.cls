VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Trigonometrics"
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
'| Class               | Trigonometrics
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
' Error constants
'
Private Const MODULE_NAME As String = "Trigonometrics"

Private Const ERR_TXT_INVALID_ARGUMENT As String = "Invalid argument"

'
' Private constants
'
Private Const PI As Double = 3.14159265358979

Private Const HALF_PI As Double = PI * 0.5

Private Const DEGREE_TO_RADIANT As Double = PI / 180#

Private Const RADIANT_TO_DEGREE As Double = 1# / DEGREE_TO_RADIANT

'
' Public methods
'

'
' Trigonometric methods
'
Public Function ArcCos(ByVal x As Double) As Double
   Select Case x
      Case Is < -1
         Err.Raise vbObjectError, MODULE_NAME, ERR_TXT_INVALID_ARGUMENT

      Case Is > 1
         Err.Raise vbObjectError, MODULE_NAME, ERR_TXT_INVALID_ARGUMENT
         
      Case -1
         ArcCos = PI

      Case 1
         ArcCos = 0#
      
      Case Else
         ArcCos = 2 * Atn(Sqr((1 - x) / (1 + x)))
   End Select
End Function

Public Function ArcSin(ByVal x As Double) As Double
   ArcSin = HALF_PI - ArcCos(x)
End Function

Public Function ArcTan2(ByVal y As Double, ByVal x As Double) As Double
   Dim result As Double

   If x = 0# Then
      If y = 0# Then
         Err.Raise vbObjectError, MODULE_NAME, ERR_TXT_INVALID_ARGUMENT
      Else
         result = HALF_PI * Sgn(y)
      End If
   Else
      result = Atn(y / x)
      
      If x < 0# Then
         If y >= 0# Then
            result = result + PI
         Else
            result = result - PI
         End If
      End If
   End If

   ArcTan2 = result
End Function

'
' Conversion methods
'
Public Function RadiantToDegree(ByVal radValue As Double) As Double
   RadiantToDegree = radValue * RADIANT_TO_DEGREE
End Function

Public Function DegreeToRadiant(ByVal degValue As Double) As Double
   DegreeToRadiant = degValue * DEGREE_TO_RADIANT
End Function
