VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AccessLockFileReader"
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
'| Class               | Access lock file reader
'|---------------------+---------------------------------------------------
'| Description         | Read access lock files
'|---------------------+---------------------------------------------------
'| Author              | Frank Schwab
'|---------------------+---------------------------------------------------
'| Version             | 1.0.1
'|---------------------+---------------------------------------------------
'| Changes             | 2020-07-21  Created. fhs
'|                     | 2020-09-14  Added comments, fixed some names. fhs
'|---------------------+---------------------------------------------------
'
Option Compare Database
Option Explicit

'
' Constants for error messages
'
Private Const ERR_STR_CLASS_NAME As String = "AccessLockFileReader"
Private Const ERR_NUM_START As Long = vbObjectError + 45122

Private Const ERR_NUM_NO_LOCK_FILE As Long = ERR_NUM_START
Private Const ERR_STR_NO_LOCK_FILE As String = "There is no lock file (neither ending with '.ldb', nor with '.laccdb')"

'
' Private constants
'
Private Const LOCK_FILE_EXTENSION_2003 As String = ".ldb"
Private Const LOCK_FILE_EXTENSION_2007 As String = ".laccdb"

'
' Private methods
'
'
'+--------------------------------------------------------------------------
'| Method           | FileExists
'|------------------+-------------------------------------------------------
'| Description      | Check if a file exists.
'|------------------+-------------------------------------------------------
'| Parameter        | filePath: Path of the file to check
'|------------------+-------------------------------------------------------
'| Return values    | True : File exists
'|                  | False: File does not exist
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2020-07-21  Created. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | ./.
'+--------------------------------------------------------------------------
'
Private Function FileExists(ByRef filePath As String) As Boolean
   FileExists = (Len(Dir$(filePath)) > 0)
End Function

'
'+--------------------------------------------------------------------------
'| Method           | GetLockFilePath
'|------------------+-------------------------------------------------------
'| Description      | Get the path of a lock file corresponding to a
'|                  | database path.
'|------------------+-------------------------------------------------------
'| Parameter        | dbPath: Path to database
'|------------------+-------------------------------------------------------
'| Return values    | Path of the lock file belonging to the supplied
'|                  | database.
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2020-07-21  Created. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | ./.
'+--------------------------------------------------------------------------
'
Private Function GetLockFilePath(ByRef dbPath As String) As String
   Dim pos As Long
   Dim withoutExtension As String
   Dim result As String

   pos = InStrRev(dbPath, ".")

   If pos = 0 Then _
      pos = Len(dbPath) + 1
   
   withoutExtension = Left$(dbPath, pos - 1)

   result = withoutExtension & LOCK_FILE_EXTENSION_2003

   If Not FileExists(result) Then
      result = withoutExtension & LOCK_FILE_EXTENSION_2007

      If Not FileExists(result) Then _
         Err.Raise ERR_NUM_NO_LOCK_FILE, ERR_STR_CLASS_NAME, ERR_STR_NO_LOCK_FILE
   End If

   GetLockFilePath = result
End Function

'
'+--------------------------------------------------------------------------
'| Method           | MyDBPath
'|------------------+-------------------------------------------------------
'| Description      | Get the path of the current database file.
'|------------------+-------------------------------------------------------
'| Parameter        | ./.
'|------------------+-------------------------------------------------------
'| Return values    | Path of the current database file name.
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2020-07-21  Created. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | ./.
'+--------------------------------------------------------------------------
'
Private Function MyDBPath() As String
   With CurrentProject
      MyDBPath = .Path & "\" & .name
   End With
End Function

'
'+--------------------------------------------------------------------------
'| Method           | StripTrailingZeroes
'|------------------+-------------------------------------------------------
'| Description      | Get the part of a string before a vbNullChar.
'|------------------+-------------------------------------------------------
'| Parameter        | aString: String with possible vbNullChars
'|------------------+-------------------------------------------------------
'| Return values    | Dictionary of all other computer names where the
'|                  | names of the other computers are the keys.
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2020-07-21  Created. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | ./.
'+--------------------------------------------------------------------------
'
Private Function StripTrailingZeroes(ByRef aString As String) As String
   Dim pos As Long

   pos = InStr(1, aString, vbNullChar)

   If pos > 0 Then
      StripTrailingZeroes = Left$(aString, pos - 1)
   Else
      StripTrailingZeroes = aString
   End If
End Function

'
'+--------------------------------------------------------------------------
'| Method           | ReadLockFileEntries
'|------------------+-------------------------------------------------------
'| Description      | Get a list of all other computers that are using the
'|                  | Access database that is named in the parameter.
'|------------------+-------------------------------------------------------
'| Parameter        | dbPath: Path of the database
'|------------------+-------------------------------------------------------
'| Return values    | Dictionary of all other computer names where the
'|                  | names of the other computers are the keys.
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2020-07-21  Created. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | ./.
'+--------------------------------------------------------------------------
'
Private Function ReadLockFileEntries(ByRef dbPath As String) As Dictionary
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

      normalizedRecord = Trim$(StripTrailingZeroes(aRecord))

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

   Set ReadLockFileEntries = result
End Function

'
' Public methods
'

'
'+--------------------------------------------------------------------------
'| Method           | GetLockFileEntriesFrom
'|------------------+-------------------------------------------------------
'| Description      | Get a list of all other computers that are using the
'|                  | Access database that is named in the parameter.
'|------------------+-------------------------------------------------------
'| Parameter        | dbPath: Path of the database
'|------------------+-------------------------------------------------------
'| Return values    | Dictionary of all other computer names where the
'|                  | names of the other computers are the keys.
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2020-07-21  Created. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | ./.
'+--------------------------------------------------------------------------
'
Public Function GetLockFileEntriesFrom(ByRef dbPath As String) As Dictionary
   Set GetLockFileEntriesFrom = ReadLockFileEntries(GetLockFilePath(dbPath))
End Function

'
'+--------------------------------------------------------------------------
'| Method           | GetMyLockFileEntries
'|------------------+-------------------------------------------------------
'| Description      | Get a list of all other computers that are using the
'|                  | current Access database.
'|------------------+-------------------------------------------------------
'| Parameter        | ./.
'|------------------+-------------------------------------------------------
'| Return values    | Dictionary of all other computer names where the
'|                  | names of the other computers are the keys.
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2020-07-21  Created. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | ./.
'+--------------------------------------------------------------------------
'
Public Function GetMyLockFileEntries() As Dictionary
   Set GetMyLockFileEntries = GetLockFileEntriesFrom(MyDBPath)
End Function

'
'+--------------------------------------------------------------------------
'| Method           | ListOfOtherComputersFrom
'|------------------+-------------------------------------------------------
'| Description      | Get a list of all other computers that are using the
'|                  | Access database that is named in the parameter.
'|------------------+-------------------------------------------------------
'| Parameter        | dbPath: Path of the database
'|------------------+-------------------------------------------------------
'| Return values    | String of all other computer names separated by
'|                  | ", ".
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2020-07-21  Created. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | ./.
'+--------------------------------------------------------------------------
'
Public Function ListOfOtherComputersFrom(ByRef dbPath As String) As String
   Dim computerList As Dictionary
   Dim aComputerName As Variant
   Dim myComputerName As String
   Dim aSI As New SystemInformation
   Dim result As String

   myComputerName = aSI.getNameOfComputer

   Set computerList = Me.GetLockFileEntriesFrom(dbPath)

   result = ""
   
   For Each aComputerName In computerList.Keys
      If aComputerName <> myComputerName Then _
         result = result & ", " & aComputerName
   Next aComputerName

   If Len(result) = 0 Then
      ListOfOtherComputersFrom = result
   Else
      ListOfOtherComputersFrom = Right$(result, Len(result) - 2)
   End If
End Function

'
'+--------------------------------------------------------------------------
'| Method           | ListMyOtherComputers
'|------------------+-------------------------------------------------------
'| Description      | Get a list of all other computers that are using the
'|                  | current Access database.
'|------------------+-------------------------------------------------------
'| Parameter        | ./.
'|------------------+-------------------------------------------------------
'| Return values    | String of all other computer names separated by
'|                  | ", ".
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2020-07-21  Created. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | ./.
'+--------------------------------------------------------------------------
'
Public Function ListMyOtherComputers() As String
   ListMyOtherComputers = ListOfOtherComputersFrom(MyDBPath)
End Function
