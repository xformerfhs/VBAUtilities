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
Private Const STR_ERR_LENGTH_INCREASED As String = "The length must not be increased without content"

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
' Public properties
'

'
'+--------------------------------------------------------------------------
'| Property         | Get blockSize
'|------------------+-------------------------------------------------------
'| Purpose          | Return current block size of this instance
'|------------------+-------------------------------------------------------
'| Parameter        | ./.
'|------------------+-------------------------------------------------------
'| Return value     | Current block size
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2018-11-29  Created. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | ./.
'|------------------+-------------------------------------------------------
'| Typical call     | aBlockSize = sb.BlockSize
'+--------------------------------------------------------------------------
Public Property Get BlockSize() As Long
   BlockSize = m_BlockSize
End Property

'
'+--------------------------------------------------------------------------
'| Property         | Let BlockSize
'|------------------+-------------------------------------------------------
'| Purpose          | Set current block size of this instance
'|------------------+-------------------------------------------------------
'| Parameter        | newBlockSize: New block size
'|------------------+-------------------------------------------------------
'| Return value     | Current block size
'|------------------+-------------------------------------------------------
'| Exceptions       | Raises exception if new block size is below minimum
'|                  | or above maximum block size.
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2018-11-29  Created. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | Minimum block size is TSBBlockSize.sbbsMinimum
'|                  | Maximum block size is TSBBlockSize.sbbsMaximum
'|------------------+-------------------------------------------------------
'| Typical call     | sb.BlockSize = newBlockSize
'+--------------------------------------------------------------------------
'
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

'
'+--------------------------------------------------------------------------
'| Property         | Get IsEmpty (read-only)
'|------------------+-------------------------------------------------------
'| Purpose          | Find out whether this instance has content, or not
'|------------------+-------------------------------------------------------
'| Parameter        | ./.
'|------------------+-------------------------------------------------------
'| Return value     | True:  No content in this instance
'|                  | False: There is content in this instance
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2018-11-29  Created. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | ./.
'|------------------+-------------------------------------------------------
'| Typical call     | if sb.IsEmpty then
'|                  |    ...    
'|                  | end if    
'+--------------------------------------------------------------------------
'
Public Property Get IsEmpty() As Boolean
   IsEmpty = (m_Length = 0)
End Property

'
'+--------------------------------------------------------------------------
'| Property         | Get content
'|------------------+-------------------------------------------------------
'| Purpose          | Return current content as string
'|------------------+-------------------------------------------------------
'| Parameter        | ./.
'|------------------+-------------------------------------------------------
'| Return value     | Current content as string
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2018-11-29  Created. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | ./.
'|------------------+-------------------------------------------------------
'| Typical call     | result = sb.Content
'+--------------------------------------------------------------------------
'
Public Property Get Content() As String
   Content = Left$(m_Content, m_Length)
End Property

'
'+--------------------------------------------------------------------------
'| Property         | Let content
'|------------------+-------------------------------------------------------
'| Purpose          | Sets current content from string
'|------------------+-------------------------------------------------------
'| Parameter        | newContent: New content
'|------------------+-------------------------------------------------------
'| Return value     | Current content as string
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2018-11-29  Created. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | Old content is deleted by this operation
'|------------------+-------------------------------------------------------
'| Typical call     | sb.Content = someString
'+--------------------------------------------------------------------------
'
Public Property Let Content(ByVal newContent As String)
   Me.Reset
   Me.Append newContent
End Property

'
'+--------------------------------------------------------------------------
'| Property         | Get length
'|------------------+-------------------------------------------------------
'| Purpose          | Return length of current content
'|------------------+-------------------------------------------------------
'| Parameter        | ./.
'|------------------+-------------------------------------------------------
'| Return value     | Length of current content
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2018-11-29  Created. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | ./.
'|------------------+-------------------------------------------------------
'| Typical call     | actLength = sb.Length
'+--------------------------------------------------------------------------
'
Public Property Get Length() As Long
   Length = m_Length
End Property

'
'+--------------------------------------------------------------------------
'| Property         | Let length
'|------------------+-------------------------------------------------------
'| Purpose          | Sets the length of the current content
'|------------------+-------------------------------------------------------
'| Parameter        | newLength: New length of content
'|------------------+-------------------------------------------------------
'| Return value     | ./-
'|------------------+-------------------------------------------------------
'| Exceptions       | Raises exception if new length is larger than
'|                  | current length or less than 0.
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2018-11-29  Created. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | The length of the content can only be shortened,
'|                  | not increased.
'|------------------+-------------------------------------------------------
'| Typical call     | sb.Length = sb.length - 10
'+--------------------------------------------------------------------------
'
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


'
' Public methods
'

'
'+--------------------------------------------------------------------------
'| Method           | SetTo
'|------------------+-------------------------------------------------------
'| Purpose          | Sets the content to be the supplied string
'|------------------+-------------------------------------------------------
'| Parameter        | text: New content of this instance
'|------------------+-------------------------------------------------------
'| Return value     | This instance
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2018-11-29  Created. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | ./.
'|------------------+-------------------------------------------------------
'| Typical call     | sb.SetTo(someString)
'+--------------------------------------------------------------------------
'Public Function SetTo(ByVal text As String) As StringBuilder
   Me.Reset
   
   Set SetTo = Me.Append(text)
End Function

'
'+--------------------------------------------------------------------------
'| Method           | Append
'|------------------+-------------------------------------------------------
'| Purpose          | Appends the supplied string to the current content
'|------------------+-------------------------------------------------------
'| Parameter        | text: text to be appended to the content of this instance
'|------------------+-------------------------------------------------------
'| Return value     | This instance
'|------------------+-------------------------------------------------------
'| Exceptions       | Raises exception if new text would make the content
'|                  | too large to hold
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2018-11-29  Created. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | ./.
'|------------------+-------------------------------------------------------
'| Typical call     | sb.Append someString
'+--------------------------------------------------------------------------
'
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

'
'+--------------------------------------------------------------------------
'| Method           | Reset
'|------------------+-------------------------------------------------------
'| Purpose          | Clears (i.e. empties) the current content
'|------------------+-------------------------------------------------------
'| Parameter        | ./-
'|------------------+-------------------------------------------------------
'| Return value     | This instance
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2018-11-29  Created. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | Reset does not change the memory space allocated to
'|                  | this instance.
'|------------------+-------------------------------------------------------
'| Typical call     | sb.Reset
'+--------------------------------------------------------------------------
'
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
