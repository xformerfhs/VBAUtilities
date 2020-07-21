VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StringBuilder"
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
'| Class               | StringBuilder
'|---------------------+---------------------------------------------------
'| Description         | A string builder with method chaining
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
Private Const MODULNAME As String = "StringBuilder"

Private Const N_START_ERROR_MESSAGE As Long = vbObjectError + 27220

Private Const ERR_BLOCK_SIZE_TOO_SMALL As Long = N_START_ERROR_MESSAGE + 1
Private Const STR_ERR_BLOCK_SIZE_TOO_SMALL As String = "Block size too small: "

Private Const ERR_BLOCK_SIZE_TOO_LARGE As Long = N_START_ERROR_MESSAGE + 2
Private Const STR_ERR_BLOCK_SIZE_TOO_LARGE As String = "Block size too large: "

Private Const ERR_TOO_LONG As Long = N_START_ERROR_MESSAGE + 3
Private Const STR_ERR_TOO_LONG As String = "Length exceeds maximum Length for strings"

Private Const ERR_LENGTH_INCREASED As Long = N_START_ERROR_MESSAGE + 4
Private Const STR_ERR_LENGTH_INCREASED As String = "The Length must not be increased without Content"

Private Const ERR_LENGTH_LESS_THAN_ZERO As Long = N_START_ERROR_MESSAGE + 5
Private Const STR_ERR_LENGTH_LESS_THAN_ZERO As String = "Length must not be less than zero"

'
' Private constants
'
Private Const DEFAULT_BLOCK_SIZE As Long = 64& * 1024&
Private Const MINIMUM_BLOCK_SIZE        As Long = 1024&
Private Const MAXIMUM_BLOCK_SIZE       As Long = 64& * 1024& * 1024&

Private Const MAXIMUM_LENGTH As Long = &H7FFFFFFF

'
' Public constants
'
Public Enum TSBBlockSize
   sbbsMinimum = MINIMUM_BLOCK_SIZE
   sbbsMaximum = MAXIMUM_BLOCK_SIZE
End Enum

'
' Instance variables
'
Private m_Content As String
Private m_Length As Long
Private m_BlockSize As Long
Private m_Size As Long

'
' Public methoden
'
Public Property Get BlockSize() As Long
   BlockSize = m_BlockSize
End Property

Public Property Let BlockSize(ByVal newBlockSize As Long)
   If newBlockSize < MINIMUM_BLOCK_SIZE Then
      Err.Raise ERR_BLOCK_SIZE_TOO_SMALL, MODULNAME, STR_ERR_BLOCK_SIZE_TOO_SMALL & Format$(newBlockSize)
   Else
      If newBlockSize > MAXIMUM_BLOCK_SIZE Then
         Err.Raise ERR_BLOCK_SIZE_TOO_LARGE, MODULNAME, STR_ERR_BLOCK_SIZE_TOO_LARGE & Format$(newBlockSize)
      Else
         m_BlockSize = newBlockSize
      End If
   End If
End Property

Public Property Get IsEmpty() As Boolean
   IsEmpty = (m_Length = 0)
End Property

Public Property Get Content() As String
   Content = Left$(m_Content, m_Length)
End Property

Public Property Let Content(ByVal newContent As String)
   Me.Reset
   Me.Append newContent
End Property

Public Property Get Length() As Long
   Length = m_Length
End Property

Public Property Let Length(ByVal newLength As Long)
   If newLength > m_Length Then
      Err.Raise ERR_LENGTH_INCREASED, MODULNAME, STR_ERR_LENGTH_INCREASED
   Else
      If newLength < 0 Then
         Err.Raise ERR_LENGTH_LESS_THAN_ZERO, MODULNAME, STR_ERR_LENGTH_LESS_THAN_ZERO
      Else
         m_Length = newLength
      End If
   End If
End Property

Public Function SetTo(ByVal text As String) As StringBuilder
   Me.Reset
   
   Set SetTo = Me.Append(text)
End Function

Public Function Append(ByVal text As String) As StringBuilder
   Dim additionalLength As Long
   
   additionalLength = Len(text)
   
   If additionalLength > 0 Then
      If (MAXIMUM_LENGTH - m_Length) < additionalLength Then
         Err.Raise ERR_TOO_LONG, MODULNAME, STR_ERR_TOO_LONG
      Else
         If (m_Length + additionalLength) > m_Size Then
            If (MAXIMUM_LENGTH - m_Length) < m_BlockSize Then
               m_BlockSize = MAXIMUM_LENGTH - m_Length
            End If
            
            m_Content = m_Content & Space$(m_BlockSize)
            m_Size = m_Size + m_BlockSize
         End If
      End If
   
      Mid$(m_Content, m_Length + 1, additionalLength) = text
   
      m_Length = m_Length + additionalLength
   End If

   Set Append = Me
End Function

Public Function Reset() As StringBuilder
   m_Length = 0
   
   Set Reset = Me
End Function

'
' Class method
'
Private Sub Class_Initialize()
   m_Length = 0
   m_BlockSize = DEFAULT_BLOCK_SIZE
   m_Size = 0
End Sub
