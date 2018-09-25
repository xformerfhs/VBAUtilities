VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AccessLockFileReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'
' Constants for error messages
'
Private Const MODULNAME As String = "AccessLockFileReader"
Private Const ERR_NUMBER_START As Long = vbObjectError + 45122

Private Const ERR_NO_LOCK_FILE As Long = ERR_NUMBER_START
Private Const STR_ERR_NO_LOCK_FILE As String = "There is no lock file (neither ending with '.ldb', nor with '.laccdb')"

'
' Private constants
'
Private Const LOCK_FILE_EXTENSION_2003 As String = ".ldb"
Private Const LOCK_FILE_EXTENSION_2007 As String = ".laccdb"

'
' Private methods
'
Private Function fileExists(ByRef dbPath As String) As Boolean
   fileExists = (Len(Dir$(dbPath)) > 0)
End Function

Private Function getLockFileName(ByRef dbPath As String) As String
   Dim pos As Long
   Dim withoutExtension As String
   Dim result As String

   pos = InStrRev(dbPath, ".")

   If pos = 0 Then _
      pos = Len(dbPath) + 1
   
   withoutExtension = Left$(dbPath, pos - 1)

   result = withoutExtension & LOCK_FILE_EXTENSION_2003

   If Not fileExists(result) Then
      result = withoutExtension & LOCK_FILE_EXTENSION_2007
      
      If Not fileExists(result) Then _
         Err.Raise ERR_NO_LOCK_FILE, MODULNAME, STR_ERR_NO_LOCK_FILE
   End If

   getLockFileName = result
End Function

Private Function myDBPath() As String
   With CurrentProject
      myDBPath = .Path & "\" & .name
   End With
End Function

Private Function stripTrailingZeroes(ByRef aString As String) As String
   Dim pos As Long

   pos = InStr(1, aString, vbNullChar)

   If pos > 0 Then
      stripTrailingZeroes = Left$(aString, pos - 1)
   Else
      stripTrailingZeroes = aString
   End If
End Function

Private Function readLockFileEntries(ByRef dbPath As String) As Dictionary
   Dim fn As Integer
   Dim aRecord As String * 32
   Dim isComputername As Boolean
   Dim result As New Dictionary
   Dim lastComputerName As String
   Dim normalizedRecord As String
   
   result.CompareMode = TextCompare

   fn = FreeFile

   Open dbPath For Binary Access Read Shared As #fn

   isComputername = True
   
   Do While Not EOF(fn)
      Get #fn, , aRecord

      normalizedRecord = Trim$(stripTrailingZeroes(aRecord))

      If Len(normalizedRecord) > 0 Then
         If isComputername Then
            lastComputerName = normalizedRecord

            If Not result.Exists(lastComputerName) Then _
               result.add Key:=normalizedRecord, item:=""
         Else
            result.item(lastComputerName) = normalizedRecord
         End If
      End If

      isComputername = Not isComputername
   Loop
   
   Close #fn

   Set readLockFileEntries = result
End Function

'
' Public methods
'
Public Function getLockFileEntriesFrom(ByRef dbPath As String) As Dictionary
   Set getLockFileEntriesFrom = readLockFileEntries(getLockFileName(dbPath))
End Function

Public Function getMyLockFileEntries() As Dictionary
   Set getMyLockFileEntries = getLockFileEntriesFrom(myDBPath)
End Function

Public Function listOfOtherComputersFrom(ByRef dbPath As String) As String
   Dim computerList As Dictionary
   Dim aComputerName As Variant
   Dim myComputerName As String
   Dim aSI As New SystemInformation
   Dim result As String

   myComputerName = aSI.getNameOfComputer

   Set computerList = Me.getLockFileEntriesFrom(dbPath)

   result = ""
   
   For Each aComputerName In computerList.Keys
      If aComputerName <> myComputerName Then _
         result = result & ", " & aComputerName
   Next aComputerName

   If Len(result) = 0 Then
      listOfOtherComputersFrom = result
   Else
      listOfOtherComputersFrom = Right$(result, Len(result) - 2)
   End If
End Function

Public Function listMyOtherComputers() As String
   listMyOtherComputers = listOfOtherComputersFrom(myDBPath)
End Function
