VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DBCompressor"
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
'| Class               | DBCompressor
'|---------------------+---------------------------------------------------
'| Description         | Compress an Access database
'|---------------------+---------------------------------------------------
'| Author              | Frank Schwab
'|---------------------+---------------------------------------------------
'| Version             | 1.0.0
'|---------------------+---------------------------------------------------
'| Changes             | 2018-09-25  Created. fhs
'|---------------------+---------------------------------------------------
'

Option Compare Database
Option Explicit

'
' Private constants for error message
'
Private Const ERR_STR_CLASS_NAME As String = "DBCompressor"
Private Const ERR_NUM_START As Long = vbObjectError + 5727

Private Const ERR_NUM_FILE_DOES_NOT_EXIST As Long = ERR_NUM_START + 0
Private Const ERR_STR_FILE_DOES_NOT_EXIST As String = "File does not exist: '"

'
' Private methods
'

'
'+--------------------------------------------------------------------------
'| Method           | DoDBCompression
'|------------------+-------------------------------------------------------
'| Description      | Create a compressed version of a database file
'|------------------+-------------------------------------------------------
'| Parameter        | oldPath: Path of the database to compress
'|                  | newPath: Path of the compressed database to create
'|------------------+-------------------------------------------------------
'| Return values    | ./.
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2018-09-25  Created. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | This method deletes the file "oldPath" and
'|                  | renames "newPath" to "oldPath". In effect it
'|                  | compresses the database "oldPath" with "newPath"
'|                  | as a temporary file.
'+--------------------------------------------------------------------------
'
Private Sub DoDBCompression(ByRef oldPath As String, ByRef newPath As String)
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

'
'+--------------------------------------------------------------------------
'| Method           | GetFilePathParts
'|------------------+-------------------------------------------------------
'| Description      | Parse a file file path and return its components
'|------------------+-------------------------------------------------------
'| Parameter        | anFSO         : A Scripting.FileSystemObject
'|                  | aFilePath     : File path to parse
'|                  | aDirectoryName: Directory of "aFilePath" (return value)
'|                  | aFileBaseName : File base name of "aFilePath" (return value)
'|                  | aFileExtension: Extension of "aFilePath" (return value)
'|------------------+-------------------------------------------------------
'| Return values    | ./.
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2018-09-25  Created. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | ./.
'+--------------------------------------------------------------------------
'
Private Sub GetFilePathParts(ByRef anFSO As Scripting.FileSystemObject, _
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

'
'+--------------------------------------------------------------------------
'| Method           | GetNewDBPath
'|------------------+-------------------------------------------------------
'| Description      | Create a random name for the temporary database
'|------------------+-------------------------------------------------------
'| Parameter        | anFSO : A Scripting.FileSystemObject
'|                  | dbPath: File path of database
'|------------------+-------------------------------------------------------
'| Return values    | ./.
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2018-09-25  Created. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | ./.
'+--------------------------------------------------------------------------
'
Private Function GetNewDBPath(ByRef anFSO As Scripting.FileSystemObject, ByRef dbPath As String) As String
   Dim dbDirectory As String
   Dim dbBaseFileName As String
   Dim dbExtension As String

   Dim anRFN As New RandomFileName

   GetFilePathParts anFSO, dbPath, dbDirectory, dbBaseFileName, dbExtension

   With anRFN
      .fileDirectory = dbDirectory
      .fileExtension = dbExtension
      .filePrefix = dbBaseFileName & "_"
      GetNewDBPath = .getUniqueRandomFilePath
   End With
End Function

'
' Public methods
'

'
'+--------------------------------------------------------------------------
'| Method           | CompressDB
'|------------------+-------------------------------------------------------
'| Description      | Compress an MS Access database
'|------------------+-------------------------------------------------------
'| Parameter        | dbPath: File path of database
'|------------------+-------------------------------------------------------
'| Return values    | ./.
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2018-09-25  Created. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | ./.
'+--------------------------------------------------------------------------
'
Public Sub CompressDB(ByRef dbPath As String)
   Dim anFSO As New Scripting.FileSystemObject

   If anFSO.FileExists(dbPath) Then
      DoDBCompression dbPath, GetNewDBPath(anFSO, dbPath)
   Else
      Err.Raise ERR_NUM_FILE_DOES_NOT_EXIST, _
                ERR_STR_CLASS_NAME, _
                ERR_STR_FILE_DOES_NOT_EXIST & dbPath & "'"
   End If
End Sub
