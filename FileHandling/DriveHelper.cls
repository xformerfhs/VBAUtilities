VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DriveHelper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'
' Error constants
'
Private Const CLASS_NAME As String = "DriveHelper"

Private Const N_START_ERROR_MESSAGE As Long = vbObjectError + 37842

Private Const ERR_NUM_UNABLE_TO_DETERMINE_DRIVE_TYPE As Long = N_START_ERROR_MESSAGE
Private Const ERR_TEXT_UNABLE_TO_DETERMINE_DRIVE_TYPE As String = "File does not exist: '"

'
' Windows API declarations
'
Private Declare Function GetDriveType _
  Lib "kernel32.dll" _
  Alias "GetDriveTypeA" ( _
  ByVal lpRootPathName As String _
  ) As Long

'
' Public constants
'
Public Enum TDriveType
   drtÚnknown = 0
   drtNoVolumeForPath = 1
   drtRemovable = 2
   drtFixed = 3
   drtRemote = 4
   drtCDROM = 5
   drtRamDisk = 6
End Enum

'
' Private methods
'
Private Sub checkDriveType(ByVal aDriveType As TDriveType)
   If aDriveType <= drtNoVolumeForPath Then _
      Err.Raise ERR_NUM_UNABLE_TO_DETERMINE_DRIVE_TYPE, _
                CLASS_NAME, _
                ERR_TEXT_UNABLE_TO_DETERMINE_DRIVE_TYPE
End Sub

'
' Public methods
'
Public Function getDriveTypeForPath(ByRef aPath As String) As TDriveType
   Dim fso As New Scripting.FileSystemObject
   
   Dim rootPath As String
   
   rootPath = fso.GetDriveName(aPath) & "\"

   Dim result As Long
   result = GetDriveType(rootPath)
   
   checkDriveType result

   getDriveTypeForPath = result
End Function

Public Function isNetworkDrive(ByRef aPath As String)
   isNetworkDrive = (getDriveTypeForPath(aPath) = drtRemote)
End Function
