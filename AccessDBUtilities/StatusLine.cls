VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StatusLine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'
' Public methods
'
Public Sub clear()
   SysCmd acSysCmdClearStatus
   
   DoEvents
End Sub

Public Sub show(ByRef aText As String)
   If Len(aText) > 0 Then
      SysCmd acSysCmdSetStatus, aText
   
      DoEvents
   Else
      Me.clear
   End If
End Sub

'
' Class methods
'
Private Sub Class_Terminate()
   Me.clear
End Sub
