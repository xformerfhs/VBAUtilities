VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "HighPrecisionTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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
