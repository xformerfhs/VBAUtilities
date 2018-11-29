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
'| BSD 3-Clause License
'|
'| Copyright (c) 2018, Frank Schwab
'|
'| All rights reserved.
'|
'| Redistribution and use in source and binary forms, with or without
'| modification, are permitted provided that the following conditions are met:
'|
'| * Redistributions of source code must retain the above copyright notice, this
'|   list of conditions and the following disclaimer.
'|
'| * Redistributions in binary form must reproduce the above copyright notice,
'|   this list of conditions and the following disclaimer in the documentation
'|   and/or other materials provided with the distribution.
'|
'| * Neither the name of the copyright holder nor the names of its
'|   contributors may be used to endorse or promote products derived from
'|   this software without specific prior written permission.
'|
'| THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS"
'| AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE
'| IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
'| DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT HOLDER OR CONTRIBUTORS BE LIABLE
'| FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL
'| DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR
'| SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER
'| CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY,
'| OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE
'| OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
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
'| Typical usage       | Dim hpt As New HighPrecisionTimer
'|                     | Dim elapsedTime As Double
'|                     | hpt.start
'|                     | ' ... Do some work ...
'|                     | elpasedTime = hpt.getElapsedTime ' Get elapsed time in seconds
'+-------------------------------------------------------------------------
'

Option Explicit

'
' External methods
'
Private Declare Function QueryPerformanceCounter Lib "kernel32" ( _
                         ByRef lpPerformanceCount As Currency) As Long

Private Declare Function QueryPerformanceFrequency Lib "kernel32" ( _
                         ByRef lpFrequency As Currency) As Long

Private m_AdjustedTickTime As Double
Private m_StartCounter As Currency

'
' Public methods
'
Public Sub start()
   QueryPerformanceCounter m_StartCounter
End Sub

Public Function getElapsedTime() As Double
   Dim currentCounter As Currency
   
   QueryPerformanceCounter currentCounter
   
   getElapsedTime = (currentCounter - m_StartCounter) * m_AdjustedTickTime
End Function

'
' Class methods
'
Private Sub Class_Initialize()
   Dim theFrequency As Currency

   QueryPerformanceFrequency theFrequency
   
   m_AdjustedTickTime = 1# / theFrequency
End Sub
