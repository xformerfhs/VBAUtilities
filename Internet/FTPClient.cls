VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FTPClient"
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
'| Class               | FTPClient
'|---------------------+---------------------------------------------------
'| Description         | Simple FTP client for VBA
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
Private Const ERROR_BASE As Long = vbObjectError + 33351
Private Const MODULE_NAME As String = "FTPClient"

Private Const ERROR_NUMBER_INVALID_VALUE As Long = ERROR_BASE + 0
Private Const ERROR_TEXT_INVALID_VALUE   As String = "Invalid value: "

Private Const ERROR_NUMBER_NO_SESSION As Long = ERROR_BASE + 1
Private Const ERROR_TEXT_NO_SESSION   As String = "No session established"

'
' Private constants
'

'
' Windows API constants
'
Private Const FILE_ATTRIBUTE_DIRECTORY As Long = &H10
Private Const FILE_ATTRIBUTE_NORMAL    As Long = &H80
Private Const FILE_ATTRIBUTE_HIDDEN    As Long = &H2

Private Const FTP_TRANSFER_TYPE_UNKNOWN As Long = 0
Private Const FTP_TRANSFER_TYPE_ASCII   As Long = 1
Private Const FTP_TRANSFER_TYPE_BINARY  As Long = 2

Private Const INTERNET_DEFAULT_FTP_PORT As Integer = 21
Private Const INTERNET_SERVICE_FTP As Long = 1
Private Const INTERNET_FLAG_PASSIVE As Long = &H8000000

Private Const INTERNET_OPEN_TYPE_PRECONFIG                   As Long = 0
Private Const INTERNET_OPEN_TYPE_DIRECT                      As Long = 1
Private Const INTERNET_OPEN_TYPE_PROXY                       As Long = 3
Private Const INTERNET_OPEN_TYPE_PRECONFIG_WITH_NO_AUTOPROXY As Long = 4

Private Const ERROR_INSUFFICIENT_BUFFER     As Long = 122
Private Const ERROR_INTERNET_EXTENDED_ERROR As Long = 12003

Private Const DEFAULT_USER_AGENT As String = "FTPClient"
Private Const DEFAULT_USER_NAME As String = "anonymous"

Private Const MAX_PATH = 260

'
' Windows API data structures
'
Private Type FILETIME
   dwLowDateTime As Long
   dwHighDateTime As Long
End Type

Private Type WIN32_FIND_DATA
   dwFileAttributes As Long
   ftCreationTime As FILETIME
   ftLastAccessTime As FILETIME
   ftLastWriteTime As FILETIME
   nFileSizeHigh As Long
   nFileSizeLow As Long
   dwReserved0 As Long
   dwReserved1 As Long
   cFileName As String * MAX_PATH
   cAlternate As String * 14
End Type

Private Type SYSTEMTIME
  wYear As Integer
  wMonth As Integer
  wDayOfWeek As Integer
  wDay As Integer
  wHour As Integer
  wMinute As Integer
  wSecond As Integer
  wMilliseconds As Integer
End Type

'
' Windows API declarations
'
Private Declare Function FileTimeToSystemTime Lib "kernel32.dll" ( _
   ByRef lpFileTime As FILETIME, _
   ByRef lpSystemTime As SYSTEMTIME) As Long

Private Declare Function InternetCloseHandle Lib "wininet.dll" (ByVal hInet As Long) As Long

Private Declare Function InternetConnect Lib "wininet.dll" Alias "InternetConnectA" ( _
   ByVal hInternetSession As Long, _
   ByVal sServerName As String, _
   ByVal nServerPort As Integer, _
   ByVal sUserName As String, _
   ByVal sPassword As String, _
   ByVal lService As Long, _
   ByVal lFlags As Long, _
   ByVal lContext As Long) As Long
   
Private Declare Function InternetOpen Lib "wininet.dll" Alias "InternetOpenA" ( _
   ByVal sAgent As String, _
   ByVal lAccessType As Long, _
   ByVal sProxyName As String, _
   ByVal sProxyBypass As String, _
   ByVal lFlags As Long) As Long
   
Private Declare Function FtpSetCurrentDirectory Lib "wininet.dll" Alias "FtpSetCurrentDirectoryA" ( _
   ByVal hFtpSession As Long, _
   ByVal lpszDirectory As String) As Long
   
Private Declare Function FtpGetCurrentDirectory Lib "wininet.dll" Alias "FtpGetCurrentDirectoryA" ( _
   ByVal hFtpSession As Long, _
   ByVal lpszCurrentDirectory As String, _
   ByRef lpdwCurrentDirectory As Long) As Long
   
Private Declare Function FtpCreateDirectory Lib "wininet.dll" Alias "FtpCreateDirectoryA" ( _
   ByVal hFtpSession As Long, _
   ByVal lpszDirectory As String) As Long
   
Private Declare Function FtpRemoveDirectory Lib "wininet.dll" Alias "FtpRemoveDirectoryA" ( _
   ByVal hFtpSession As Long, _
   ByVal lpszDirectory As String) As Long
   
Private Declare Function FtpDeleteFile Lib "wininet.dll" Alias "FtpDeleteFileA" ( _
   ByVal hFtpSession As Long, _
   ByVal lpszFileName As String) As Long
   
Private Declare Function FtpRenameFile Lib "wininet.dll" Alias "FtpRenameFileA" ( _
   ByVal hFtpSession As Long, _
   ByVal lpszExisting As String, _
   ByVal lpszNew As String) As Long
   
Private Declare Function FtpGetFile Lib "wininet.dll" Alias "FtpGetFileA" ( _
   ByVal hConnect As Long, _
   ByVal lpszRemoteFile As String, _
   ByVal lpszNewFile As String, _
   ByVal fFailIfExists As Integer, _
   ByVal m_FlagsAndAttributes As Long, _
   ByVal m_Flags As Long, _
   ByRef dwContext As Long) As Long

Private Declare Function FtpPutFile Lib "wininet.dll" Alias "FtpPutFileA" ( _
   ByVal hConnect As Long, _
   ByVal lpszLocalFile As String, _
   ByVal lpszNewRemoteFile As String, _
   ByVal m_Flags As Long, _
   ByVal dwContext As Long) As Long

Private Declare Function InternetGetLastResponseInfo Lib "wininet.dll" Alias "InternetGetLastResponseInfoA" ( _
   ByRef lpdwError As Long, _
   ByVal lpszBuffer As String, _
   ByRef lpdwBufferLength As Long) As Long
   
Private Declare Function FtpFindFirstFile Lib "wininet.dll" Alias "FtpFindFirstFileA" ( _
   ByVal hFtpSession As Long, _
   ByVal lpszSearchFile As String, _
   ByRef lpFindFileData As WIN32_FIND_DATA, _
   ByVal m_Flags As Long, _
   ByVal dwContent As Long) As Long
   
Private Declare Function InternetFindNextFile Lib "wininet.dll" Alias "InternetFindNextFileA" ( _
   ByVal hFind As Long, _
   ByRef lpvFindData As WIN32_FIND_DATA) As Long

'
' Instance variables
'
Private m_UserAgent As String

Private m_Session As Long
Private m_InternetHandle As Long

Private m_ErrorCode As Long
Private m_ErrorText As String

Private m_MM As MessageManager

'
' Private methods
'
Private Function GetMessageManager() As MessageManager
   If m_MM Is Nothing Then
      Set m_MM = New MessageManager
   End If

   Set GetMessageManager = m_MM
End Function

Private Sub RaiseInvalidValueException(ByRef errorDescription As String)
   Err.Raise ERROR_NUMBER_INVALID_VALUE, _
             MODULE_NAME, _
             ERROR_TEXT_INVALID_VALUE & _
                errorDescription
End Sub

Private Function GetInternetErrorMessage() As String
   Dim errorCode As Long
   Dim ErrorMessage As String
   Dim bufferLength As Long

   bufferLength = 0
   InternetGetLastResponseInfo errorCode, vbNullString, bufferLength

   ErrorMessage = Space$(bufferLength)
   InternetGetLastResponseInfo errorCode, ErrorMessage, bufferLength

   GetInternetErrorMessage = ErrorMessage
End Function

Private Function GetErrorCodeFromFTPMessage(ByRef ftpMessage As String) As Long
   Dim lines() As String
   Dim i As Integer
   Dim pos As Integer
   Dim messageNumber As Long

   GetErrorCodeFromFTPMessage = 0

   lines = Split(ftpMessage, vbCrLf)

   On Error GoTo ErrorHandler
   
   For i = LBound(lines) To UBound(lines)
      pos = InStr(1, lines(i), " ")

      If pos = 4 Then
         messageNumber = CLng(Left$(lines(i), pos - 1))

         If messageNumber >= 400 Then _
            GetErrorCodeFromFTPMessage = messageNumber
      End If
   Next i

ExitFunction:
   Exit Function

ErrorHandler:
   Resume ExitFunction
End Function

Private Function GetAPIErrorMessage(ByRef functionName As String, ByVal errorCode As Long) As String
   GetAPIErrorMessage = GetMessageManager.GetMessageForModuleErrorCode(errorCode, "wininet.dll")
End Function

Private Sub HandleAPIError(ByRef functionName As String, ByVal errorCode As Long)
   If errorCode <> ERROR_INTERNET_EXTENDED_ERROR Then
      m_ErrorCode = errorCode
      m_ErrorText = GetAPIErrorMessage(functionName, errorCode)
   Else
      m_ErrorText = GetInternetErrorMessage
      m_ErrorCode = GetErrorCodeFromFTPMessage(m_ErrorText)
   End If
End Sub

Private Sub ClearError()
   m_ErrorCode = 0
   m_ErrorText = ""
End Sub

Private Sub CheckSession()
   If Not Me.IsSessionEstablished Then _
      Err.Raise ERROR_NUMBER_NO_SESSION, _
                MODULE_NAME, _
                ERROR_TEXT_NO_SESSION
End Sub

Private Sub EnsureEnvironment()
   CheckSession
   ClearError
End Sub

Private Function ConvertFileTimeToDateTime(ByRef aFileTime As FILETIME) As Date
   Dim systemTimeEntry As SYSTEMTIME

   If (aFileTime.dwLowDateTime Or aFileTime.dwHighDateTime) <> 0 Then
      FileTimeToSystemTime aFileTime, systemTimeEntry
   
      With systemTimeEntry
         ConvertFileTimeToDateTime = DateSerial(.wYear, .wMonth, .wDay) + _
                                     TimeSerial(.wHour, .wMinute, .wSecond)
      End With
   Else
      ConvertFileTimeToDateTime = 0#
   End If
End Function


Private Function GetInternetHandle() As Long
   If m_InternetHandle = 0 Then
      m_InternetHandle = InternetOpen(m_UserAgent, INTERNET_OPEN_TYPE_PRECONFIG, vbNullString, vbNullString, 0&)

      If m_InternetHandle = 0 Then _
         HandleAPIError "InternetOpen", Err.LastDllError
   End If

   GetInternetHandle = m_InternetHandle
End Function

Private Sub CloseInternetHandle()
   If m_InternetHandle <> 0 Then
      InternetCloseHandle m_InternetHandle
      m_InternetHandle = 0
   End If
End Sub

'
' Public properties
'
Public Property Get IsSessionEstablished() As Boolean
   IsSessionEstablished = (m_Session <> 0)
End Property

Public Property Get HadError() As Boolean
   HadError = (Len(m_ErrorText) > 0)
End Property

Public Property Get errorCode() As Long
   errorCode = m_ErrorCode
End Property

Public Property Get ErrorMessage() As String
   ErrorMessage = m_ErrorText
End Property

'
' Public methods
'
Public Function Connect(ByRef serverName As String, _
                        Optional ByVal portNumber As Integer = INTERNET_DEFAULT_FTP_PORT, _
                        Optional ByVal usePassiveFTP As Boolean = True, _
                        Optional ByRef userName As String = DEFAULT_USER_NAME, _
                        Optional ByRef password As String = DEFAULT_USER_NAME) As Boolean
   ClearError

   Me.Disconnect

   Connect = False

   m_Session = InternetConnect(GetInternetHandle(), _
                               serverName, _
                               portNumber, _
                               userName, _
                               password, _
                               INTERNET_SERVICE_FTP, _
                               IIf(usePassiveFTP, INTERNET_FLAG_PASSIVE, 0&), _
                               0&)

   If m_Session <> 0& Then
      Connect = True
   Else
      HandleAPIError "InternetConnect", Err.LastDllError
   End If
End Function
   
Public Sub Disconnect()
   If IsSessionEstablished Then
      InternetCloseHandle m_Session
      m_Session = 0
   End If
End Sub
   
Public Function GetFile(ByRef remoteFileName As String, _
                        ByRef localFileName As String, _
                        Optional ByVal isBinaryTransfer As Boolean = True) As Boolean
   EnsureEnvironment

   GetFile = False

   If FtpGetFile(m_Session, _
                 remoteFileName, _
                 localFileName, _
                 0, _
                 FILE_ATTRIBUTE_NORMAL, _
                 IIf(isBinaryTransfer, FTP_TRANSFER_TYPE_BINARY, FTP_TRANSFER_TYPE_ASCII), _
                 0&) <> 0& Then
         GetFile = True
   Else
      HandleAPIError "FtpGetFile", Err.LastDllError
   End If
End Function

Public Function PutFile(ByRef localFileName As String, _
                        ByRef remoteFileName As String, _
                        Optional ByVal isBinaryTransfer As Boolean = True) As Boolean
   EnsureEnvironment

   PutFile = False

   If FtpPutFile(m_Session, _
                 localFileName, _
                 remoteFileName, _
                 IIf(isBinaryTransfer, FTP_TRANSFER_TYPE_BINARY, FTP_TRANSFER_TYPE_ASCII), _
                 0&) <> 0& Then
      PutFile = True
   Else
      HandleAPIError "FtpPutFile", Err.LastDllError
   End If
End Function

Public Function CreateDirectory(ByRef directoryName As String) As Boolean
   EnsureEnvironment

   CreateDirectory = False

   If FtpCreateDirectory(m_Session, directoryName) <> 0& Then
      CreateDirectory = True
   Else
      HandleAPIError "FtpCreateDirectory", Err.LastDllError
   End If
End Function

Public Function RemoveDirectory(ByRef directoryName As String) As Boolean
   EnsureEnvironment

   RemoveDirectory = False

   If FtpRemoveDirectory(m_Session, directoryName) <> 0& Then
      RemoveDirectory = False
   Else
      HandleAPIError "FtpRemoveDirectory", Err.LastDllError
   End If
End Function

Public Function RenameFile(ByRef oldName As String, ByRef newName As String) As Boolean
   EnsureEnvironment

   RenameFile = False

   If FtpRenameFile(m_Session, oldName, newName) <> 0& Then
      RenameFile = True
   Else
      HandleAPIError "FtpRenameFile", Err.LastDllError
   End If
End Function

Public Function DeleteFile(ByRef filename As String) As Boolean
   EnsureEnvironment

   DeleteFile = False

   If FtpDeleteFile(m_Session, filename) <> 0& Then
      DeleteFile = True
   Else
      HandleAPIError "FtpDeleteFile", Err.LastDllError
   End If
End Function

Public Function GetCurrentDirectory() As String
   Dim currentDirectory As String
   Dim directoryNameLength As Long
   Dim lastErrorCode As Long

   EnsureEnvironment

   If FtpGetCurrentDirectory(m_Session, vbNullString, directoryNameLength) = 0 Then
      lastErrorCode = Err.LastDllError

      If lastErrorCode <> ERROR_INSUFFICIENT_BUFFER Then _
         HandleAPIError "FtpGetCurrentDirectory-GetDataLength", Err.LastDllError
   End If

   currentDirectory = Space$(directoryNameLength)

   If FtpGetCurrentDirectory(m_Session, currentDirectory, directoryNameLength) = 0& Then _
      HandleAPIError "FtpGetCurrentDirectory-GetData", Err.LastDllError

   GetCurrentDirectory = Left$(currentDirectory, directoryNameLength)
End Function

Public Function SetCurrentDirectory(ByRef directoryName As String) As Boolean
   EnsureEnvironment

   SetCurrentDirectory = False

   If FtpSetCurrentDirectory(m_Session, directoryName) <> 0& Then
      SetCurrentDirectory = True
   Else
      HandleAPIError "FtpSetCurrentDirectory", Err.LastDllError
   End If
End Function

Public Function DirFiles(Optional ByRef fileFilter As String = "*.*") As Scripting.Dictionary
   Dim findData As WIN32_FIND_DATA
   Dim findHandle As Long
   Dim result As New Scripting.Dictionary
   Dim rc As Long
   Dim directoryEntry As DirectoryInformation
   Dim systemTimeEntry As SYSTEMTIME

   EnsureEnvironment

   findHandle = FtpFindFirstFile(m_Session, fileFilter, findData, 0&, 0&)

   If findHandle <> 0& Then
      Do
         Set directoryEntry = New DirectoryInformation

         With directoryEntry
            .Name = Left$(findData.cFileName, InStr(1, findData.cFileName, vbNullChar, vbBinaryCompare) - 1)
            .IsDirectory = ((findData.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) <> 0)
            .CreationTime = ConvertFileTimeToDateTime(findData.ftCreationTime)
            .LastAccessTime = ConvertFileTimeToDateTime(findData.ftLastAccessTime)
            .LastWriteTime = ConvertFileTimeToDateTime(findData.ftLastWriteTime)
            .FileSizeHigh = findData.nFileSizeHigh
            .FileSizeLow = findData.nFileSizeLow
         End With
         
         result.Add Key:=directoryEntry.Name, Item:=directoryEntry

         rc = InternetFindNextFile(findHandle, findData)
      Loop While rc <> 0

      InternetCloseHandle findHandle
   Else
      HandleAPIError "FtpFindFirstFile", Err.LastDllError

      If Me.errorCode = 450 Then
         ClearError
      End If
   End If

   Set DirFiles = result
End Function

'
' Class methods
'
Private Sub Class_Initialize()
   m_UserAgent = DEFAULT_USER_AGENT
   
   m_InternetHandle = 0
   m_Session = 0

   m_ErrorCode = 0
   m_ErrorText = ""
End Sub

Private Sub Class_Terminate()
   Disconnect
   CloseInternetHandle
End Sub
