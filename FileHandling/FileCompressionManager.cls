VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FileCompressionManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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
''' Constants for GetVolumeInformation
'''
Private Const FILE_CASE_PRESERVED_NAMES         As Long = &H2&
Private Const FILE_CASE_SENSITIVE_SEARCH        As Long = &H1&
Private Const FILE_FILE_COMPRESSION             As Long = &H10&
Private Const FILE_NAMED_STREAMS                As Long = &H40000
Private Const FILE_PERSISTENT_ACLS              As Long = &H8&
Private Const FILE_READ_ONLY_VOLUME             As Long = &H80000
Private Const FILE_SEQUENTIAL_WRITE_ONCE        As Long = &H100000
Private Const FILE_SUPPORTS_ENCRYPTION          As Long = &H20000
Private Const FILE_SUPPORTS_EXTENDED_ATTRIBUTES As Long = &H800000
Private Const FILE_SUPPORTS_OBJECT_IDS          As Long = &H10000
Private Const FILE_SUPPORTS_OPEN_BY_FILE_ID     As Long = &H1000000
Private Const FILE_SUPPORTS_REPARSE_POINTS      As Long = &H80&
Private Const FILE_SUPPORTS_SPARSE_FILES        As Long = &H40&
Private Const FILE_SUPPORTS_TRANSACTIONS        As Long = &H200000
Private Const FILE_SUPPORTS_USN_JOURNAL         As Long = &H2000000
Private Const FILE_UNICODE_ON_DISK              As Long = &H4&
Private Const FILE_VOLUME_IS_COMPRESSED         As Long = &H8000
Private Const FILE_VOLUME_QUOTAS                As Long = &H20&

'''
''' Constants for dwDesiredAccess
'''
Private Const GENERIC_ALL     As Long = &H10000000
Private Const GENERIC_READ    As Long = &H80000000
Private Const GENERIC_WRITE   As Long = &H40000000
Private Const GENERIC_EXECUTE As Long = &H20000000

'''
''' Constants for dwShareMode
'''
Private Const FILE_SHARE_DELETE  As Long = &H4
Private Const FILE_SHARE_NOTHING As Long = &H0
Private Const FILE_SHARE_READ    As Long = &H1
Private Const FILE_SHARE_WRITE   As Long = &H2

'''
''' Constants for dwCreationDisposition
'''
Private Const CREATE_ALWAYS     As Long = 2
Private Const CREATE_NEW        As Long = 1
Private Const OPEN_ALWAYS       As Long = 4
Private Const OPEN_EXISTING     As Long = 3
Private Const TRUNCATE_EXISTING As Long = 5
 
'''
''' Constants for dwFlagsAndAttributes
'''
Private Const FILE_ATTRIBUTE_ARCHIVE       As Long = &H20
Private Const FILE_ATTRIBUTE_ENCRYPTED     As Long = &H4000&
Private Const FILE_ATTRIBUTE_HIDDEN        As Long = &H2
Private Const FILE_ATTRIBUTE_NORMAL        As Long = &H80
Private Const FILE_ATTRIBUTE_OFFLINE       As Long = &H1000&
Private Const FILE_ATTRIBUTE_READONLY      As Long = &H1
Private Const FILE_ATTRIBUTE_SYSTEM        As Long = &H4
Private Const FILE_ATTRIBUTE_TEMPORARY     As Long = &H100&
Private Const FILE_FLAG_DELETE_ON_CLOSE    As Long = &H4000000
Private Const FILE_FLAG_NO_BUFFERING       As Long = &H20000000
Private Const FILE_FLAG_OPEN_NO_RECALL     As Long = &H100000
Private Const FILE_FLAG_OPEN_REPARSE_POINT As Long = &H200000
Private Const FILE_FLAG_OVERLAPPED         As Long = &H40000000
Private Const FILE_FLAG_POSIX_SEMANTICS    As Long = &H1000000
Private Const FILE_FLAG_RANDOM_ACCESS      As Long = &H10000000
Private Const FILE_FLAG_SEQUENTIAL_SCAN    As Long = &H8000000
Private Const FILE_FLAG_SESSION_AWARE      As Long = &H800000
Private Const FILE_FLAG_WRITE_THROUGH      As Long = &H80000000

Private Const INVALID_HANDLE_VALUE As Long = -1&

'''
''' Constants for ioctl
'''
Private Const FILE_ANY_ACCESS   As Long = &H0&
Private Const FILE_READ_ACCESS  As Long = &H1&
Private Const FILE_WRITE_ACCESS As Long = &H2&

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

Private Function DoesFilesystemSupportCompression(ByRef drivename As String) As Boolean
   DoesFilesystemSupportCompression = ((GetVolumeFlags(drivename) And FILE_FILE_COMPRESSION) <> 0)
End Function

Private Function GetDriveName(ByRef fileName As String) As String
   GetDriveName = m_FSO.GetDriveName(m_FSO.GetAbsolutePathName(fileName)) & "\"
End Function

Private Function MakeIOCtlCode(ByVal deviceType As Long, ByVal accessType As Long, ByVal functionCode As Long, ByVal ioMethod As Long) As Long
   MakeIOCtlCode = ((deviceType * 65536) Or _
                    (accessType * 16384&) Or _
                    (functionCode * 4&) Or _
                    ioMethod)
End Function

Private Function GetCompressionAttribute(ByRef handle As Long) As Boolean
   Dim controlParameter As Integer
   Dim returnedBytesCount As Long

   Dim rc As Long

   GetCompressionAttribute = False

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

   rc = DeviceIoControl(handle, _
                        MakeIOCtlCode(FILE_DEVICE_FILE_SYSTEM, FILE_READ_ACCESS Or FILE_WRITE_ACCESS, IOCTL_FUNCTION_SET_COMPRESSION, METHOD_BUFFERED), _
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
   Dim handle As Long
   
   handle = CreateFile(fileName, _
                       GENERIC_READ Or GENERIC_WRITE, _
                       FILE_SHARE_READ, _
                       0, _
                       creationDisposition, _
                       FILE_ATTRIBUTE_NORMAL, _
                       0)

   If handle <> INVALID_HANDLE_VALUE Then
      If DoesFilesystemSupportCompression(GetDriveName(fileName)) Then _
         SetCompressionAttribute handle, shouldCompress

      If CloseHandle(handle) = 0 Then _
         HandleAPIError "CloseHandle", fileName
   Else
      HandleAPIError "CreateFile", fileName
   End If
End Sub

Private Function GetCompressionAttributeForFile(ByRef fileName As String) As Boolean
   Dim handle As Long
   
   GetCompressionAttributeForFile = False

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
