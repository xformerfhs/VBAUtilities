VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "HighPrecisionTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
'+-------------------------------------------------------------------------
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
'| Class               | HighPrecisionTimer
'|---------------------+---------------------------------------------------
'| Description         | Implements a high precision timer by using the
'|                     | the QueryPerformanceCounter API
'|---------------------+---------------------------------------------------
'| Author              | Frank Schwab
'|---------------------+---------------------------------------------------
'| Version             | 1.0.0
'|---------------------+---------------------------------------------------
'| Changes             | 2018-11-29  Created. fhs
'|---------------------+---------------------------------------------------
'| Remarks             | The "QueryPerformanceCounter" and "QueryPerformanceFrequency"
'|                     | API calls return 8 byte integers. VBA has no 8 byte
'|                     | integer data type. So we use the VBA data type "Currency"
'|                     | which implements 8 byte integers scaled by the
'|                     | factor 10,000. This data type is used in such a way
'|                     | that in the calculation of the elpased time this
'|                     | scale factor cancels out.
'|---------------------+---------------------------------------------------
'| Typical usage       | Dim hpt As New HighPrecisionTimer
'|                     | Dim elapsedTime As Double
'|                     | hpt.start
'|                     | ' ... Do some work ...
'|                     | elpasedTime = hpt.getElapsedTime ' Get elapsed time in seconds
'+-------------------------------------------------------------------------
'

Option Explicit

'
' Windows API declarations
'
Private Declare Function QueryPerformanceCounter Lib "kernel32" ( _
                         ByRef lpPerformanceCount As Currency) As Long

Private Declare Function QueryPerformanceFrequency Lib "kernel32" ( _
                         ByRef lpFrequency As Currency) As Long

Private m_ScaledTickTime As Double
Private m_StartCounter   As Currency

'
' Public methods
'

'
'+--------------------------------------------------------------------------
'| Method           | Start
'|------------------+-------------------------------------------------------
'| Purpose          | Starts the high precision timer
'|------------------+-------------------------------------------------------
'| Parameter        | ./.
'|------------------+-------------------------------------------------------
'| Return value     | ./.
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2018-11-29  Created. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | ./.
'|------------------+-------------------------------------------------------
'| Typical call     | See class comment.
'+--------------------------------------------------------------------------
'
Public Sub Start()
   QueryPerformanceCounter m_StartCounter
End Sub

'
'+--------------------------------------------------------------------------
'| Method           | GetElapsedTime
'|------------------+-------------------------------------------------------
'| Purpose          | Get the time that has passed in seconds since
'|                  | calling the "start" method
'|------------------+-------------------------------------------------------
'| Parameter        | ./.
'|------------------+-------------------------------------------------------
'| Return value     | Time in seconds since calling "start"
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2018-11-29  Created. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | ./.
'|------------------+-------------------------------------------------------
'| Typical call     | See class comment.
'+--------------------------------------------------------------------------
' 
Public Function GetElapsedTime() As Double
   Dim currentCounter As Currency
   
   QueryPerformanceCounter currentCounter

   ' Here the scale factor 10,000 of the counters and the scaled tick
    ' time just cancel out each other.
   getElapsedTime = (currentCounter - m_StartCounter) * m_ScaledTickTime
End Function

'
' Class methods
'
Private Sub Class_Initialize()
   Dim theFrequency As Currency

   QueryPerformanceFrequency theFrequency
   
   ' Here we calculate the length of a tick scaled by the factor 10,000
   ' of the "Currency" data type.
   m_ScaledTickTime = 1# / theFrequency
End Sub
