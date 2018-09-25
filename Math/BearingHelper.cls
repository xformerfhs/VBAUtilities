VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BearingHelper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'
' Public methods
'
Public Function calculateInBetweenBearing(ByVal bearingFrom As Integer, ByVal bearingTo As Integer, ByVal fraction As Double) As Integer
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

   calculateInBetweenBearing = result
End Function

Public Function calculateMeanBearing(ByVal bearingFrom As Integer, ByVal bearingTo As Integer) As Integer
   calculateMeanBearing = calculateInBetweenBearing(bearingFrom, bearingTo, 0.5)
End Function
