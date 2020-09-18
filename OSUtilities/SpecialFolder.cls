VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SpecialFolder"
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
'| Class               | SpecialFolder
'|---------------------+---------------------------------------------------
'| Description         | Get special folder names
'|---------------------+---------------------------------------------------
'| Author              | Frank Schwab
'|---------------------+---------------------------------------------------
'| Version             | 1.0.0
'|---------------------+---------------------------------------------------
'| Changes             | 2020-07-23  Created. fhs
'|---------------------+---------------------------------------------------
'
Option Explicit

'
' Error handling constants
'
Private Const ERROR_BASE As Long = vbObjectError + 512
Private Const MODULE_NAME As String = "SpecialFolder"

Private Const ERROR_NUMBER_API_ERROR As Long = ERROR_BASE + 0
Private Const ERROR_TEXT_API_ERROR As String = "Error "

'
' Windows API constants
'
Private Const CSIDL_PERSONAL       As Long = 5
Private Const CSIDL_APPDATA        As Long = &H1A
Private Const CSIDL_COMMON_APPDATA As Long = &H23

Private Const SHGFP_TYPE_CURRENT As Long = 0
Private Const SHGFP_TYPE_DEFAULT As Long = 1

Private Const S_OK         As Long = 0
Private Const S_FALSE      As Long = 1
Private Const E_INVALIDARG As Long = &H80070057

'
' Windows API declarations
'

Private Declare Function SHGetFolderPath _
Lib "shell32.dll" _
Alias "SHGetFolderPathA" ( _
  ByVal hwnd As Long, _
  ByVal csidl As Long, _
  ByVal hToken As Long, _
  ByVal dwFlags As Long, _
  ByVal pszPath As String) As Long

'
' Instance variables
'
Private m_Path As String

'
' Private methods
'
Private Sub HandleAPIError(ByVal errorCode As Long, ByRef source As String, Optional ByRef additionalInformation As String = "")
   If errorCode = S_FALSE Then
      Err.Raise ERROR_NUMBER_API_ERROR, _
                MODULE_NAME, _
                source & " is a virtual folder"
   Else
      Err.Raise ERROR_NUMBER_API_ERROR, _
                MODULE_NAME, _
                source & " is an invalid folder"
   End If
End Sub

Private Function MakeStringFromNullTerminatedString(ByRef aNullTerminatedString As String) As String
   Dim pos As Long
   
   pos = InStr(1, aNullTerminatedString, vbNullChar)
   
   If pos > 0 Then
      MakeStringFromNullTerminatedString = Left$(aNullTerminatedString, pos - 1)
   Else
      MakeStringFromNullTerminatedString = aNullTerminatedString
   End If
End Function

'
' Public methods
'
Public Function GetDefaultAppDataPath() As String
   Dim result As Long

   result = SHGetFolderPath(0, CSIDL_APPDATA, 0, SHGFP_TYPE_DEFAULT, m_Path)

   If result = S_OK Then
      GetDefaultAppDataPath = MakeStringFromNullTerminatedString(m_Path)
   Else
      HandleAPIError result, "CSIDL_APPDATA-Default"
   End If
End Function

Public Function GetCurrentAppDataPath() As String
   Dim result As Long

   result = SHGetFolderPath(0, CSIDL_APPDATA, 0, SHGFP_TYPE_CURRENT, m_Path)

   If result = S_OK Then
      GetCurrentAppDataPath = MakeStringFromNullTerminatedString(m_Path)
   Else
      HandleAPIError result, "CSIDL_APPDATA-Current"
   End If
End Function

Public Function GetDefaultCommonAppDataPath() As String
   Dim result As Long

   result = SHGetFolderPath(0, CSIDL_COMMON_APPDATA, 0, SHGFP_TYPE_DEFAULT, m_Path)

   If result = S_OK Then
      GetDefaultCommonAppDataPath = MakeStringFromNullTerminatedString(m_Path)
   Else
      HandleAPIError result, "CSIDL_COMMON_APPDATA-Default"
   End If
End Function

Public Function GetCurrentCommonAppDataPath() As String
   Dim result As Long

   result = SHGetFolderPath(0, CSIDL_COMMON_APPDATA, 0, SHGFP_TYPE_CURRENT, m_Path)

   If result = S_OK Then
      GetCurrentCommonAppDataPath = MakeStringFromNullTerminatedString(m_Path)
   Else
      HandleAPIError result, "CSIDL_COMMON_APPDATA-Current"
   End If
End Function

Public Function GetDefaultPersonalPath() As String
   Dim result As Long

   result = SHGetFolderPath(0, CSIDL_PERSONAL, 0, SHGFP_TYPE_DEFAULT, m_Path)

   If result = S_OK Then
      GetDefaultPersonalPath = MakeStringFromNullTerminatedString(m_Path)
   Else
      HandleAPIError result, "CSIDL_PERSONAL-Default"
   End If
End Function

Public Function GetCurrentPersonalPath() As String
   Dim result As Long

   result = SHGetFolderPath(0, CSIDL_PERSONAL, 0, SHGFP_TYPE_CURRENT, m_Path)

   If result = S_OK Then
      GetCurrentPersonalPath = MakeStringFromNullTerminatedString(m_Path)
   Else
      HandleAPIError result, "CSIDL_PERSONAL-Current"
   End If
End Function

'
' Class methods
'
Private Sub Class_Initialize()
   m_Path = Space$(260)  ' 260 = MAX_PATH
End Sub
