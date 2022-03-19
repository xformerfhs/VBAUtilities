VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BinaryFileInputStream"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
'+-------------------------------------------------------------------------
'|
'| SPDX-FileCopyrightText: 2022 Frank Schwab
'|
'| SPDX-License-Identifier: MIT
'|
'| Copyright 2022, Frank Schwab
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
'| Class               | BinaryFileInputStream
'|---------------------+---------------------------------------------------
'| Description         | Binary file input stream
'|---------------------+---------------------------------------------------
'| Author              | Frank Schwab
'|---------------------+---------------------------------------------------
'| Version             | 1.0.0
'|---------------------+---------------------------------------------------
'| Changes             | 2022-03-19  Created. fhs
'|---------------------+---------------------------------------------------
'| Typical usage       | Dim buffer(1 To 32768) As Byte
'|                     | Dim bfis As New BinaryFileInputStream
'|                     | bfis.ForFile filePath
'|                     | Dim readCount As Long
'|                     | Do
'|                     |    readCount = bfis.Read(buffer)
'|                     |    If readCount > 0 Then
'|                     |       ' Do something with the bytes read
'|                     |    Else
'|                     |       Exit Do
'|                     |    End If
'|                     | Loop
'|                     | bfis.CloseStream()
'+---------------------+---------------------------------------------------
'

Option Explicit

'
' Private constants
'

'
' Constants for error messages
'
Private Const ERR_STR_SOURCE   As String = "BinaryFileInputStream"
Private Const ERR_BASE As Long = vbObjectError + 57031

' Invalid buffer size
Private Const ERR_NUM_INVALID_BUFFER_SIZE As Long = ERR_BASE
Private Const ERR_STR_INVALID_BUFFER_SIZE As String = "invalid buffer size"

' Stream not initialized
Private Const ERR_NUM_NOT_INITIALIZED As Long = ERR_BASE + 1
Private Const ERR_STR_NOT_INITIALIZED As String = "stream not initialized"

'
' Constants for class function
'
Private Const MIN_BUFFER_SIZE As Long = 512
Private Const MAX_BUFFER_SIZE As Long = 64& * 1024& * 1024&

'
' Instance variables
'

Private m_Fn As Integer
Private m_FileLength As Long

Private m_LastPosition As Long

'
' Private methods
'

'
'+--------------------------------------------------------------------------
'| Method           | CheckState
'|------------------+-------------------------------------------------------
'| Description      | Check state of stream reader
'|------------------+-------------------------------------------------------
'| Parameter        | ./.
'|------------------+-------------------------------------------------------
'| Return values    | ./.
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2022-03-19  Created. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | This method throws an exception if the stream is closed
'+--------------------------------------------------------------------------
'
Private Sub CheckState()
   If m_Fn = 0 Then _
      Err.Raise ERR_NUM_INVALID_BUFFER_SIZE, _
                ERR_STR_SOURCE, _
                ERR_STR_INVALID_BUFFER_SIZE
End Sub

'
' Public methods
'

'
'+--------------------------------------------------------------------------
'| Method           | ForFile
'|------------------+-------------------------------------------------------
'| Description      | Open the stream for a file
'|------------------+-------------------------------------------------------
'| Parameter        | filePath: Path of file
'|------------------+-------------------------------------------------------
'| Return values    | ./.
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2022-03-19  Created. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | ./.
'+--------------------------------------------------------------------------
'
Public Sub ForFile(ByRef filePath As String)
   m_Fn = FreeFile
   Open filePath For Binary Access Read Lock Write As #m_Fn

   m_FileLength = LOF(m_Fn)
   
   m_LastPosition = 0
End Sub

'
'+--------------------------------------------------------------------------
'| Method           | Read
'|------------------+-------------------------------------------------------
'| Description      | Read data from stream
'|------------------+-------------------------------------------------------
'| Parameter        | buffer: Buffer to place the data into
'|------------------+-------------------------------------------------------
'| Return values    | Number of bytes read
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2022-03-19  Created. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | ./.
'+--------------------------------------------------------------------------
'
Public Function Read(ByRef buffer() As Byte) As Long
   CheckState
   
   Dim bufferSize As Long
   bufferSize = UBound(buffer) - LBound(buffer) + 1

   If (bufferSize < MIN_BUFFER_SIZE) Or (bufferSize > MAX_BUFFER_SIZE) Then _
      Err.Raise ERR_NUM_INVALID_BUFFER_SIZE, _
                ERR_STR_SOURCE, _
                ERR_STR_INVALID_BUFFER_SIZE

   Dim result As Long

   If Not EOF(m_Fn) Then
      Get #m_Fn, , buffer

      If m_LastPosition >= (m_FileLength - bufferSize) Then
         result = m_FileLength - m_LastPosition
      Else
         result = bufferSize
      End If

      m_LastPosition = m_LastPosition + result
   Else
      result = 0
   End If

   Read = result
End Function

'
'+--------------------------------------------------------------------------
'| Method           | CloseStream
'|------------------+-------------------------------------------------------
'| Description      | Close the stream
'|------------------+-------------------------------------------------------
'| Parameter        | ./.
'|------------------+-------------------------------------------------------
'| Return values    | ./.
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2022-03-19  Created. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | The file streamed from is closed
'+--------------------------------------------------------------------------
'
Public Sub CloseStream()
   If m_Fn <> 0 Then
      Close #m_Fn
      m_Fn = 0
   End If
End Sub

'
' Class methods
'
Private Sub Class_Terminate()
   CloseStream
End Sub
