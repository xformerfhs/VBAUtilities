VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RandomFileName"
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
'| Class               | RandomFileName
'|---------------------+---------------------------------------------------
'| Description         | Generates random file names
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
' Private constants for error messages
'
Private Const CLASS_NAME As String = "RandomFileName"

Private Const N_START_ERROR_MESSAGE As Long = vbObjectError + 1470

Private Const ERR_NUM_INVALID_DIRECTORY        As Long = N_START_ERROR_MESSAGE + 1
Private Const ERR_NUM_DIRECTORY_DOES_NOT_EXIST As Long = N_START_ERROR_MESSAGE + 2

Private Const ERR_TEXT_INVALID_DIRECTORY        As String = "Specified directory is invalid"
Private Const ERR_TEXT_DIRECTORY_DOES_NOT_EXIST As String = "Specified directory does not exist: "

'
' Private constants
'
Private Const VALID_FILE_NAME_CHARACTERS As String = "abcdefghijklmnopqrstuvwxyz0123456789^°!§$%&()=´`+~#',;-_@"

Private Const SEPARATOR_EXTENSION        As String = "."
Private Const SEPARATOR_PATH             As String = "\"
Private Const SEPARATOR_PATH_ALTERNATIVE As String = "/"

Private Const FILENAME_UNDERSCORE As String = "_"

'
' Instance variables
'
Private m_Directory As String
Private m_Extension As String

'
' Attributes
'
Public filePrefix As String
Public fileSuffix As String

'
' Private methods
'
Private Function GetFileExtensionWithSeparator() As String
   If Len(m_Extension) > 0 Then
      GetFileExtensionWithSeparator = SEPARATOR_EXTENSION & m_Extension
   Else
      GetFileExtensionWithSeparator = m_Extension
   End If
End Function

Private Function GetFileDirectoryWithSeparator() As String
   GetFileDirectoryWithSeparator = m_Directory & SEPARATOR_PATH
End Function

Private Function RandomLongInRange(ByVal fromValue As Long, ByVal toValue As Long) As Long
   RandomLongInRange = Int((toValue - fromValue + 1) * Rnd) + fromValue
End Function

Private Function GetRandomFileNameCharacter() As String
   GetRandomFileNameCharacter = Mid$(VALID_FILE_NAME_CHARACTERS, RandomLongInRange(1, Len(VALID_FILE_NAME_CHARACTERS)), 1)
End Function

Private Function CreateRandomFileName() As String
   Dim result As String
   Dim i As Integer
   
   result = Space$(8)

   For i = 1 To Len(result)
      Mid$(result, i, 1) = GetRandomFileNameCharacter
   Next i
   
   CreateRandomFileName = result
End Function

Private Function ConstructFileName() As String
   Dim result As String

   If Len(filePrefix) > 0 Then _
      result = filePrefix & FILENAME_UNDERSCORE

   result = result & CreateRandomFileName

   If Len(fileSuffix) > 0 Then _
      result = result & FILENAME_UNDERSCORE & fileSuffix

   ConstructFileName = result & GetFileExtensionWithSeparator
End Function

Private Function CreateUniqueRandomFilePath()
   Dim filePathProbe As String

   Do
      filePathProbe = GetFileDirectoryWithSeparator & ConstructFileName

      If Len(Dir$(filePathProbe)) = 0 Then
         CreateUniqueRandomFilePath = filePathProbe
         Exit Do
      End If
   Loop
End Function

'
' Public attributes
'
Public Property Get FileDirectory() As String
   FileDirectory = m_Directory
End Property

Public Property Let FileDirectory(ByRef newDirectory As String)
   Dim normalizedDirectory As String
   Dim lastCharacter As String * 1
   
   normalizedDirectory = Trim$(newDirectory)

   If Len(normalizedDirectory) = 0 Then
      Err.Raise ERR_NUM_INVALID_DIRECTORY, CLASS_NAME, ERR_TEXT_INVALID_DIRECTORY
   Else
      If Len(Dir$(normalizedDirectory, vbDirectory)) = 0 Then
         Err.Raise ERR_NUM_DIRECTORY_DOES_NOT_EXIST, _
                   CLASS_NAME, _
                   ERR_TEXT_DIRECTORY_DOES_NOT_EXIST & normalizedDirectory
      Else
         lastCharacter = Right$(normalizedDirectory, 1)

         If lastCharacter <> SEPARATOR_PATH And _
            lastCharacter <> SEPARATOR_PATH_ALTERNATIVE Then
            m_Directory = normalizedDirectory
         Else
            m_Directory = Left$(normalizedDirectory, Len(normalizedDirectory) - 1)
         End If
      End If
   End If
End Property

Public Property Get FileExtension() As String
   FileExtension = m_Extension
End Property

Public Property Let FileExtension(ByRef newExtension As String)
   Dim normalizedExtension As String
   Dim length As Long
   Dim pos As Long

   normalizedExtension = Trim$(newExtension)

   length = Len(normalizedExtension)

   If length > 0 Then
      pos = InStrRev(normalizedExtension, SEPARATOR_EXTENSION)

      If pos = 0 Then
         m_Extension = normalizedExtension
      Else
         m_Extension = Right$(normalizedExtension, length - pos)
      End If
   Else
      m_Extension = normalizedExtension
   End If
End Property

'
' Public methods
'
Public Function GetUniqueRandomFilePath() As String
   GetUniqueRandomFilePath = CreateUniqueRandomFilePath
End Function

'
' Class methods
'
Private Sub Class_Initialize()
   Randomize

   m_Directory = ".\"
   m_Extension = ""

   filePrefix = ""
   fileSuffix = ""
End Sub
