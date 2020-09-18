VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DriveHelper"
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
'|
'|-------------------------------------------------------------------------
'| Class               | DriveHelper
'|---------------------+---------------------------------------------------
'| Description         | Get information about a drive
'|---------------------+---------------------------------------------------
'| Author              | Frank Schwab
'|---------------------+---------------------------------------------------
'| Version             | 1.0.1
'|---------------------+---------------------------------------------------
'| Changes             | 2020-07-20  Created. fhs
'|                     | 2020-08-30  Made 64 bit compatible. fhs
'|---------------------+---------------------------------------------------
'

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
Private Declare PtrSafe Function GetDriveType _
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
Private Sub CheckDriveType(ByVal aDriveType As TDriveType)
   If aDriveType <= drtNoVolumeForPath Then _
      Err.Raise ERR_NUM_UNABLE_TO_DETERMINE_DRIVE_TYPE, _
                CLASS_NAME, _
                ERR_TEXT_UNABLE_TO_DETERMINE_DRIVE_TYPE
End Sub

'
' Public methods
'
Public Function GetDriveTypeForPath(ByRef aPath As String) As TDriveType
   Dim fso As New Scripting.FileSystemObject
   
   Dim rootPath As String
   
   rootPath = fso.GetDriveName(aPath) & "\"

   Dim result As Long
   result = GetDriveType(rootPath)
   
   CheckDriveType result

   GetDriveTypeForPath = result
End Function

Public Function IsNetworkDrive(ByRef aPath As String)
   IsNetworkDrive = (GetDriveTypeForPath(aPath) = drtRemote)
End Function
