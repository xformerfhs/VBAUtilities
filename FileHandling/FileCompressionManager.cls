VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FileCompressionManager"
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
'|
'|-------------------------------------------------------------------------
'| Class               | FileCompressionManager
'|---------------------+---------------------------------------------------
'| Description         | Manage compression state of a file
'|---------------------+---------------------------------------------------
'| Author              | Frank Schwab
'|---------------------+---------------------------------------------------
'| Version             | 1.0.0
'|---------------------+---------------------------------------------------
'| Changes             | 2020-07-20  Created. fhs
'|---------------------+---------------------------------------------------
'

Option Explicit


'
' Error handling constants
'
Private Const ERROR_BASE As Long = vbObjectError + 3866
Private Const MODULE_NAME As String = "FileCompressionManager"

Private Const ERROR_NUMBER_API_ERROR As Long = ERROR_BASE + 0
Private Const ERROR_TEXT_API_ERROR As String = "Error "

'
' Private constants
'

'''
''' Relevant constant for GetVolumeInformation
'''
Private Const FILE_FILE_COMPRESSION As Long = &H10&

'''
''' Constants for dwDesiredAccess
'''
Private Const GENERIC_READ    As Long = &H80000000
Private Const GENERIC_WRITE   As Long = &H40000000

'''
''' Constants for dwShareMode
'''
Private Const FILE_SHARE_READ    As Long = &H1

'''
''' Constants for dwCreationDisposition
'''
Private Const CREATE_ALWAYS As Long = 2
Private Const OPEN_EXISTING As Long = 3

'''
''' Constant for dwFlagsAndAttributes
'''
Private Const FILE_ATTRIBUTE_NORMAL As Long = &H80

'''
''' Constant for invalid handle
'''
Private Const INVALID_HANDLE_VALUE As Long = -1&

'''
''' Constants for ioctl
'''
Private Const FILE_ANY_ACCESS As Long = &H0&
Private Const FILE_READ_DATA  As Long = &H1&
Private Const FILE_WRITE_DATA As Long = &H2&

Private Const FILE_DEVICE_FILE_SYSTEM As Long = &H9&

Private Const METHOD_BUFFERED As Long = 0&

Private Const IOCTL_FUNCTION_GET_COMPRESSION As Long = 15&
Private Const IOCTL_FUNCTION_SET_COMPRESSION As Long = 16&

Private Const COMPRESSION_FORMAT_NONE    As Long = 0&
Private Const COMPRESSION_FORMAT_DEFAULT As Long = 1&

'
' Declare Windows API function
'
Private Declare Function GetVolumeInformation Lib "kernel32.dll" Alias "GetVolumeInformationA" ( _
   ByVal lpRootPathName As String, _
   ByVal lpVolumeNameBuffer As LongPtr, _
   ByVal nVolumeNameSize As Long, _
   ByVal lpVolumeSerialNumber As LongPtr, _
   ByVal lpMaximumComponentLength As LongPtr, _
   ByVal lpFileSystemFlags As LongPtr, _
   ByVal lpFileSystemNameBuffer As LongPtr, _
   ByVal nFileSystemNameSize As Long) As Long

Private Declare Function CreateFile Lib "kernel32.dll" Alias "CreateFileA" ( _
   ByVal lpFileName As String, _
   ByVal dwDesiredAccess As Long, _
   ByVal dwShareMode As Long, _
   ByVal lpSecurityAttributes As Long, _
   ByVal dwCreationDisposition As Long, _
   ByVal dwFlagsAndAttributes As Long, _
   ByVal hTemplateFile As Long) As Long

Private Declare Function CloseHandle Lib "kernel32.dll" ( _
   ByVal hObject As Long) As Long

Private Declare Function DeviceIoControl Lib "kernel32.dll" ( _
   ByVal hDevice As Long, _
   ByVal dwIoControlCode As Long, _
   ByVal lpInBuffer As LongPtr, _
   ByVal nInBufferSize As Long, _
   ByVal lpOutBuffer As LongPtr, _
   ByVal nOutBufferSize As Long, _
   ByVal lpBytesReturned As Long, _
   ByVal lpOverlapped As Long) As Long

'
' Instance variables
'
Private m_FSO As New Scripting.FileSystemObject

'
' Private methods
'
Private Function FileExists(ByRef fileName As String) As Boolean
   FileExists = (Len(Dir$(fileName)) > 0)
End Function

Private Sub HandleAPIError(ByRef source As String, Optional ByRef additionalInformation As String = "")
   Dim ErrorCode As Long

   ErrorCode = Err.LastDllError

   Dim mm As New MessageManager

   Dim errorText As String
   
   errorText = ERROR_TEXT_API_ERROR & _
               Format$(ErrorCode) & _
               " (0x" & _
               Hex$(ErrorCode) & _
               ") on call to '" & _
               source & _
               "'"
   
   If Len(additionalInformation) <> 0 Then _
      errorText = errorText & " (" & additionalInformation & ")"

   Err.Raise ERROR_NUMBER_API_ERROR, _
             MODULE_NAME, _
             errorText & _
                ": " & _
                mm.GetMessageForWindowsErrorCode(ErrorCode)
End Sub

Private Function GetVolumeFlags(ByRef drivename As String) As Long
   Dim result As Long

   GetVolumeFlags = 0

   If GetVolumeInformation(drivename, 0, 0, 0, 0, VarPtr(result), 0, 0) <> 0 Then
      GetVolumeFlags = result
   Else
      HandleAPIError "GetVolumeInformation"
   End If
End Function

Private Function GetDriveName(ByRef fileName As String) As String
   GetDriveName = m_FSO.GetDriveName(m_FSO.GetAbsolutePathName(fileName)) & "\"
End Function

Private Function DoesFilesystemSupportCompression(ByRef fileName As String) As Boolean
   DoesFilesystemSupportCompression = ((GetVolumeFlags(GetDriveName(fileName)) And FILE_FILE_COMPRESSION) <> 0)
End Function

Private Function MakeIOCtlCode(ByVal deviceType As Long, ByVal accessType As Long, ByVal functionCode As Long, ByVal ioMethod As Long) As Long
   Dim result As Long

   result = ((deviceType * &H10000) Or _
            (accessType * &H4000&) Or _
            (functionCode * 4&) Or _
            ioMethod)

   MakeIOCtlCode = result
End Function

Private Function GetCompressionAttribute(ByRef handle As Long) As Boolean
   Dim controlParameter As Integer
   Dim returnedBytesCount As Long

   Dim rc As Long

   GetCompressionAttribute = False

   '
   ' The "required access" parameter *must* be set to "FILE_ANY_ACCESS" in order to read the compression state.
   ' Setting it to "FILE_READ_DATA" results in an "invalid parameter" error.
   '
   ' From winioctl.h: #define FSCTL_GET_COMPRESSION CTL_CODE(FILE_DEVICE_FILE_SYSTEM, 15, METHOD_BUFFERED, FILE_ANY_ACCESS)
   '
   rc = DeviceIoControl(handle, _
                        MakeIOCtlCode(FILE_DEVICE_FILE_SYSTEM, FILE_ANY_ACCESS, IOCTL_FUNCTION_GET_COMPRESSION, METHOD_BUFFERED), _
                        0&, _
                        0&, _
                        VarPtr(controlParameter), _
                        LenB(controlParameter), _
                        VarPtr(returnedBytesCount), _
                        0&)

   If rc <> 0 Then
      GetCompressionAttribute = (controlParameter <> COMPRESSION_FORMAT_NONE)
   Else
      HandleAPIError "DeviceIoControl(FSCTL_GET_COMPRESSION)"
   End If
End Function

Private Sub SetCompressionAttribute(ByRef handle As Long, ByVal shouldCompress As Boolean)
   Dim controlParameter As Integer
   Dim returnedBytesCount As Long

   Dim rc As Long

   If shouldCompress Then
      controlParameter = COMPRESSION_FORMAT_DEFAULT
   Else
      controlParameter = COMPRESSION_FORMAT_NONE
   End If

   '
   ' From winioctl.h: #define FSCTL_SET_COMPRESSION CTL_CODE(FILE_DEVICE_FILE_SYSTEM, 16, METHOD_BUFFERED, FILE_READ_DATA | FILE_WRITE_DATA)
   '
   rc = DeviceIoControl(handle, _
                        MakeIOCtlCode(FILE_DEVICE_FILE_SYSTEM, FILE_READ_DATA Or FILE_WRITE_DATA, IOCTL_FUNCTION_SET_COMPRESSION, METHOD_BUFFERED), _
                        VarPtr(controlParameter), _
                        LenB(controlParameter), _
                        0&, _
                        0&, _
                        VarPtr(returnedBytesCount), _
                        0&)

   If rc = 0 Then _
      HandleAPIError "DeviceIoControl(FSCTL_SET_COMPRESSION)"
End Sub

Private Sub OpenFileAndSetCompressionState(ByRef fileName As String, ByVal creationDisposition As Long, ByVal shouldCompress As Boolean)
   If DoesFilesystemSupportCompression(fileName) Then
      Dim handle As Long

'
' It is absolutely necessary to specify "GENERIC_READ or GENERIC_WRITE" as the
' desired access parameter even though the attribute is only meant to be set
' and not read. However, setting the compression attribute implies to read
' the current compression attribute first.
'
      handle = CreateFile(fileName, _
                          GENERIC_READ Or GENERIC_WRITE, _
                          FILE_SHARE_READ, _
                          0, _
                          creationDisposition, _
                          FILE_ATTRIBUTE_NORMAL, _
                          0)

      If handle <> INVALID_HANDLE_VALUE Then
            SetCompressionAttribute handle, shouldCompress

         If CloseHandle(handle) = 0 Then _
            HandleAPIError "CloseHandle", fileName
      Else
         HandleAPIError "CreateFile", fileName
      End If
   End If
End Sub

Private Function GetCompressionAttributeForFile(ByRef fileName As String) As Boolean
   GetCompressionAttributeForFile = False

   If DoesFilesystemSupportCompression(fileName) Then
      Dim handle As Long
   
      handle = CreateFile(fileName, _
                          GENERIC_READ, _
                          FILE_SHARE_READ, _
                          0, _
                          OPEN_EXISTING, _
                          FILE_ATTRIBUTE_NORMAL, _
                          0)

      If handle <> INVALID_HANDLE_VALUE Then
         GetCompressionAttributeForFile = GetCompressionAttribute(handle)

         If CloseHandle(handle) = 0 Then _
            HandleAPIError "CloseHandle", fileName
      Else
         HandleAPIError "CreateFile", fileName
      End If
   End If
End Function


'
' Public methods
'
Public Sub CreateCompressedFile(ByRef fileName As String)
   OpenFileAndSetCompressionState fileName, CREATE_ALWAYS, True
End Sub

Public Sub SetCompression(ByRef fileName As String)
   OpenFileAndSetCompressionState fileName, OPEN_EXISTING, True
End Sub

Public Sub ClearCompression(ByRef fileName As String)
   OpenFileAndSetCompressionState fileName, OPEN_EXISTING, False
End Sub

Public Function IsFileCompressed(ByRef fileName As String) As Boolean
   IsFileCompressed = GetCompressionAttributeForFile(fileName)
End Function
