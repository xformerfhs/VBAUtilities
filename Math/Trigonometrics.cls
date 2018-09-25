VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Trigonometrics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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
