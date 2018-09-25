VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DBCompressor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'
' Private constants for error message
'
Private Const CLASS_NAME As String = "DBCompressor"

Private Const N_START_ERROR_MESSAGE As Long = vbObjectError + 5727

Private Const ERR_NUM_FILE_DOES_NOT_EXIST As Long = N_START_ERROR_MESSAGE + 0

Private Const ERR_TEXT_FILE_DOES_NOT_EXIST As String = "File does not exist: '"

'
' Private methods
'
Private Sub doDBCompression(ByRef oldPath As String, ByRef newPath As String)
   '
   ' Compress source db to temp db
   '
   DBEngine.CompactDatabase oldPath, newPath

   '
   ' Delete source db and rename compressed db to source db
   '
   Kill oldPath
   Name newPath As oldPath
End Sub

Private Sub getFilePathParts(ByRef anFSO As Scripting.FileSystemObject, _
                             ByRef aFilePath As String, _
                             ByRef aDirectoryName As String, _
                             ByRef aFileBaseName As String, _
                             ByRef aFileExtension As String)
   With anFSO
      aDirectoryName = .GetParentFolderName(aFilePath)
      aFileBaseName = .GetBaseName(aFilePath)
      aFileExtension = .GetExtensionName(aFilePath)
   End With
End Sub

Private Function getNewDBPath(ByRef anFSO As Scripting.FileSystemObject, ByRef dbPath As String) As String
   Dim dbDirectory As String
   Dim dbBaseFileName As String
   Dim dbExtension As String

   Dim anRFN As New RandomFileName

   getFilePathParts anFSO, dbPath, dbDirectory, dbBaseFileName, dbExtension

   With anRFN
      .fileDirectory = dbDirectory
      .fileExtension = dbExtension
      .filePrefix = dbBaseFileName & "_"
      getNewDBPath = .getUniqueRandomFilePath
   End With
End Function

'
' Public methods
'
Public Sub compressDB(ByRef dbPath As String)
   Dim anFSO As New Scripting.FileSystemObject

   If anFSO.FileExists(dbPath) Then
      doDBCompression dbPath, getNewDBPath(anFSO, dbPath)
   Else
      Err.Raise ERR_NUM_FILE_DOES_NOT_EXIST, _
                CLASS_NAME, _
                ERR_TEXT_FILE_DOES_NOT_EXIST & dbPath & "'"
   End If
End Sub
