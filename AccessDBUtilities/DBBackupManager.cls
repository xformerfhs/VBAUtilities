VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DBBackupManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'
' Private methods
'
Private Function getDateIndexedFileList(ByRef fileList As Scripting.Files) As Variant
   Dim result As New Scripting.Dictionary
   result.CompareMode = BinaryCompare

   Dim aFile As Scripting.File
   
   For Each aFile In fileList
      result.Add Key:=aFile.DateCreated, Item:=aFile
   Next
   
   Set getDateIndexedFileList = result
End Function

Private Function getDateSortedFileList(ByRef fileList As Scripting.Files) As Scripting.File()
   Dim aFile As Scripting.File

   Dim dateIndexedFileList As Scripting.Dictionary
   
   Set dateIndexedFileList = getDateIndexedFileList(fileList)

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
   
   getDateSortedFileList = result
End Function

Private Sub deleteOldBackups(ByRef fso As Scripting.FileSystemObject, ByRef backupDirectory As String, ByVal noOfFilesToKeep As Integer)
   Dim backupFolder As Scripting.Folder

   Set backupFolder = fso.GetFolder(backupDirectory)
   
   Dim fileList As Scripting.Files

   Set fileList = backupFolder.Files
   
   If fileList.count > noOfFilesToKeep Then
      Dim dateSortedFileList() As Scripting.File

      dateSortedFileList = getDateSortedFileList(fileList)

      Dim i As Integer

      For i = LBound(dateSortedFileList) To UBound(dateSortedFileList) - noOfFilesToKeep
         dateSortedFileList(i).Delete
      Next i
   End If
End Sub

'
' Public methods
'
Public Sub backupDB(ByRef dbPath As String, ByVal noOfFilesToKeep As Integer)
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

   deleteOldBackups fso, backupDirectory, noOfFilesToKeep
   
   MsgBox "Database backuped to file '" & backupPath & "'", vbInformation Or vbOKOnly, "Database backuped"
End Sub
