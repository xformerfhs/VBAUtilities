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
'| Class               | HighPrecisionTimer
'|---------------------+---------------------------------------------------
'| Description         | A timer that uses the high precision
'|                     | QueryPerformanceCounter Windows API
'|---------------------+---------------------------------------------------
'| Author              | Frank Schwab
'|---------------------+---------------------------------------------------
'| Version             | 1.0.1
'|---------------------+---------------------------------------------------
'| Changes             | 2006-04-03  Created. fhs
'|                     | 2020-08-24  Made 64 bit compatible. fhs
'|---------------------+---------------------------------------------------
'| Remarks             | This class uses the "Currency" data type to
'|                     | store the performance counter. This data type
'|                     | uses 8 bytes of storage and implicitely has
'|                     | 4 decimal places.
'|                     |
'|                     | However the performance counter is just an 8 byte
'|                     | integer with no decimal places. So, why is the
'|                     | "Currency" data type used here?
'|                     |
'|                     | Well, this is because the 4 decimal places cancel
'|                     | out. The elapsed time is calculated by the following
'|                     | formula:
'|                     |
'|                     | elapsedTime = (PerformanceCounter#Now - PerformanceCounter#Start) / PerformanceFrequency
'|                     |
'|                     | All of the participating variables have the 4 places
'|                     | shift. So the formula with "Currency" variables
'|                     | really reads:
'|                     |
'|                     | elapsedTime = ((PerformanceCounter#Now * 0.0001) - (PerformanceCounter#Start * 0.0001)) / (PerformanceFrequency * 0.0001)
'|                     |
'|                     | One can rewrite this as:
'|                     |
'|                     | elapsedTime = ((PerformanceCounter#Now - PerformanceCounter#Start) * 0.0001) / (PerformanceFrequency * 0.0001)
'|                     |
'|                     | Now you can see that the "0.0001" factors just cancel
'|                     | each other.
'|                     |
'|                     | So, using "Currency" is a convenient way to store
'|                     | 8 byte integers and calculate divisions.
'|---------------------+---------------------------------------------------
'| Typical usage       | Dim hpc As New HighPrecisionTimer
'|                     | Dim elapsedTime As Double
'|                     | ..
'|                     | hpc.Start
'|                     | ' Do some work here
'|                     | elapsedTime = hpc.GetElapsedTime
'+-------------------------------------------------------------------------

Option Compare Binary
Option Explicit

'
' API declarations
'
Private Declare PtrSafe Function QueryPerformanceCounter Lib "kernel32" ( _
                         ByRef lpPerformanceCount As Currency) As Long

Private Declare PtrSafe Function QueryPerformanceFrequency Lib "kernel32" ( _
                         ByRef lpFrequency As Currency) As Long


'
' Instance variables
'
Private m_TimePerTick  As Double
Private m_StartCounter As Currency


'
' Public methods
'
Public Sub Start()
   QueryPerformanceCounter m_StartCounter
End Sub

Public Function GetElapsedTime() As Double
   Dim currentCounter As Currency
   
   QueryPerformanceCounter currentCounter
   
   GetElapsedTime = (currentCounter - m_StartCounter) * m_TimePerTick
End Function

'
' Class methods
'
Private Sub Class_Initialize()
   Dim theFrequency As Currency

   QueryPerformanceFrequency theFrequency
   
   m_TimePerTick = 1# / theFrequency
End Sub
