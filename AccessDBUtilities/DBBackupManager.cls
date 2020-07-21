VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DBBackupManager"
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
'|-------------------------------------------------------------------------
'| Class               | DBBackupManager
'|---------------------+---------------------------------------------------
'| Description         | Manage database backups
'|---------------------+---------------------------------------------------
'| Author              | Frank Schwab
'|---------------------+---------------------------------------------------
'| Version             | 1.0.0
'|---------------------+---------------------------------------------------
'| Changes             | 2020-07-20  Created. fhs
'|---------------------+---------------------------------------------------
'

Option Compare Database
Option Explicit

'
' Private methods
'
Private Function GetDateIndexedFileList(ByRef fileList As Scripting.Files) As Variant
   Dim result As New Scripting.Dictionary
   result.CompareMode = BinaryCompare

   Dim aFile As Scripting.File
   
   For Each aFile In fileList
      result.Add Key:=aFile.DateCreated, Item:=aFile
   Next
   
   Set GetDateIndexedFileList = result
End Function

Private Function GetDateSortedFileList(ByRef fileList As Scripting.Files) As Scripting.File()
   Dim aFile As Scripting.File

   Dim dateIndexedFileList As Scripting.Dictionary
   
   Set dateIndexedFileList = GetDateIndexedFileList(fileList)

   Dim dh As New DictionaryHelper
   
   Dim dateList() As Variant
   
   dateList = dh.getKeysAsArray(dateIndexedFileList)
   
   Dim aSorter As New Sorter
   
   aSorter.insertionSort dateList

   Dim result() As Scripting.File

   ReDim result(LBound(dateList) To UBound(dateList))
   
   Dim i As Integer
   
   For i = LBound(dateList) To UBound(dateList)
      Set result(i) = dateIndexedFileList.Item(dateList(i))
   Next i
   
   GetDateSortedFileList = result
End Function

Private Sub DeleteOldBackups(ByRef fso As Scripting.FileSystemObject, ByRef backupDirectory As String, ByVal noOfFilesToKeep As Integer)
   Dim backupFolder As Scripting.Folder

   Set backupFolder = fso.GetFolder(backupDirectory)
   
   Dim fileList As Scripting.Files

   Set fileList = backupFolder.Files
   
   If fileList.count > noOfFilesToKeep Then
      Dim dateSortedFileList() As Scripting.File

      dateSortedFileList = GetDateSortedFileList(fileList)

      Dim i As Integer

      For i = LBound(dateSortedFileList) To UBound(dateSortedFileList) - noOfFilesToKeep
         dateSortedFileList(i).Delete
      Next i
   End If
End Sub

'
' Public methods
'
Public Sub BackupDB(ByRef dbPath As String, ByVal noOfFilesToKeep As Integer)
   Dim fso As New Scripting.FileSystemObject

   Dim dbDirectory As String

   dbDirectory = fso.GetParentFolderName(dbPath)
   
   Dim backupDirectory As String
   backupDirectory = fso.GetParentFolderName(dbPath) & "\Backups"
   
   If Not fso.FolderExists(backupDirectory) Then _
      fso.CreateFolder backupDirectory

   Dim backupPath As String

   backupPath = backupDirectory & "\" & fso.GetBaseName(dbPath) & "_" & Format$(Now, "YYYY\-MM\-DDTHH\-NN\-SS") & "." & fso.GetExtensionName(dbPath)

   fso.CopyFile dbPath, backupPath

   DeleteOldBackups fso, backupDirectory, noOfFilesToKeep
   
   MsgBox "Database backuped to file '" & backupPath & "'", vbInformation Or vbOKOnly, "Database backuped"
End Sub
