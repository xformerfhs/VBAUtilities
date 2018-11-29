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
'| BSD 3-Clause License
'|
'| Copyright (c) 2018, Frank Schwab
'|
'| All rights reserved.
'|
'| Redistribution and use in source and binary forms, with or without
'| modification, are permitted provided that the following conditions are met:
'|
'| * Redistributions of source code must retain the above copyright notice, this
'|   list of conditions and the following disclaimer.
'|
'| * Redistributions in binary form must reproduce the above copyright notice,
'|   this list of conditions and the following disclaimer in the documentation
'|   and/or other materials provided with the distribution.
'|
'| * Neither the name of the copyright holder nor the names of its
'|   contributors may be used to endorse or promote products derived from
'|   this software without specific prior written permission.
'|
'| THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS"
'| AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE
'| IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
'| DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT HOLDER OR CONTRIBUTORS BE LIABLE
'| FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL
'| DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR
'| SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER
'| CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY,
'| OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE
'| OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
'|
'|-------------------------------------------------------------------------
'| Class               | StringBuilder
'|---------------------+---------------------------------------------------
'| Description         | Implements a class to build strings dynamically
'|                     | without copying them again and again when text is
'|                     | appended.
'|---------------------+---------------------------------------------------
'| Author              | Frank Schwab
'|---------------------+---------------------------------------------------
'| Version             | 1.0.0
'|---------------------+---------------------------------------------------
'| Changes             | 2018-11-29  Created. fhs
'|---------------------+---------------------------------------------------
'| Remarks             | Preallocating strings in large blocks and not copying 
'|                     | them again and again when text is appended is much faster
'|                     | than VBA's standard behavior.
'|                     |
'|                     | Default block size is 64 KiB which is quite large and
'|                     | can be adjusted just after the instance as been created.
'|                     |
'|                     | A class instance always returns itself as the
'|                     | result of an operation so that operations can be
'|                     | concatenated.
'|---------------------+---------------------------------------------------
'| Typical usage       | Dim sb As New StringBuilder
'|                     | sb.blockSize = 10000
'|                     | Dim finalString As String
'|                     | finalString = sb.setTo(someString).append(otherString).append(stillOtherString).content
'+-------------------------------------------------------------------------
'

Option Explicit

'
' Private constants for error messages
'
Private Const MODULNAME As String = "StringBuilder"

Private Const N_START_ERROR_MESSAGE As Long = vbObjectError + 27220

Private Const ERR_BLOCK_SIZE_TOO_SMALL As Long = N_START_ERROR_MESSAGE
Private Const STR_ERR_BLOCK_SIZE_TOO_SMALL As String = "Block size too small: "

Private Const ERR_BLOCK_SIZE_TOO_LARGE As Long = N_START_ERROR_MESSAGE + 1
Private Const STR_ERR_BLOCK_SIZE_TOO_LARGE As String = "Block size too large: "

Private Const ERR_TOO_LONG As Long = N_START_ERROR_MESSAGE + 2
Private Const STR_ERR_TOO_LONG As String = "Length exceeds maximum length for strings"

Private Const ERR_LENGTH_INCREASED As Long = N_START_ERROR_MESSAGE + 3
Private Const STR_ERR_LENGTH_INCREASED As String = "The length must not be increased without content"

Private Const ERR_LENGTH_LESS_THAN_ZERO As Long = N_START_ERROR_MESSAGE + 4
Private Const STR_ERR_LENGTH_LESS_THAN_ZERO As String = "Length must not be less than zero"

'
' Private constants
'
Private Const DEFAULT_BLOCK_SIZE As Long = 64& * 1024&
Private Const MINIMUM_BLOCK_SIZE As Long = 1024&
Private Const MAXIMUM_BLOCK_SIZE As Long = 64& * 1024& * 1024&

Private Const MAXIMUM_LENGTH As Long = &H7FFFFFFF

'
' Public constants
'
Public Enum TSBBlockSize
  sbbsMinimum  = MINIMUM_BLOCK_SIZE
  sbbsgMaximum = MAXIMUM_BLOCK_SIZE
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
'| Typical call     | aBlockSize = sb.blockSize
'+--------------------------------------------------------------------------
'
Public Property Get blockSize() As Long
   blockSize = m_BlockSize
End Property

'
'+--------------------------------------------------------------------------
'| Property         | Let blockSize
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
'| Typical call     | sb.blockSize = newBlockSize
'+--------------------------------------------------------------------------
'
Public Property Let blockSize(ByVal newBlockSize As Long)
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
'| Typical call     | result = sb.content
'+--------------------------------------------------------------------------
'
Public Property Get content() As String
   content = Left$(m_Content, m_Length)
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
'| Typical call     | sb.content = someString
'+--------------------------------------------------------------------------
'
Public Property Let content(ByVal newContent As String)
   Me.reset
   Me.append newContent
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
'| Typical call     | actLength = sb.length
'+--------------------------------------------------------------------------
'
Public Property Get length() As Long
   length = m_Length
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
'| Typical call     | sb.length = sb.length - 10
'+--------------------------------------------------------------------------
'
Public Property Let length(ByVal newLength As Long)
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
'+--------------------------------------------------------------------------
'| Property         | Get isEmpty (read-only)
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
'| Typical call     | if sb.isEmpty then
'|                  |    ...    
'|                  | end if    
'+--------------------------------------------------------------------------
'
Public Property Get isEmpty() As Boolean
   isEmpty = (m_Length = 0)
End Property

'
' Public methods
'

'
'+--------------------------------------------------------------------------
'| Method           | setTo
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
'| Typical call     | sb.setTo(someString)
'+--------------------------------------------------------------------------
'
Public Function setTo(ByVal text As String) As StringBuilder
   Me.reset
   
   Set setTo = Me.append(text)
End Function

'
'+--------------------------------------------------------------------------
'| Method           | append
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
'| Typical call     | sb.append someString
'+--------------------------------------------------------------------------
'
Public Function append(ByVal text As String) As StringBuilder
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

   Set append = Me
End Function

'
'+--------------------------------------------------------------------------
'| Method           | reset
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
'| Typical call     | sb.reset
'+--------------------------------------------------------------------------
'
Public Function reset() As StringBuilder
   m_Length = 0
   
   Set reset = Me
End Function

'
' Class methods
'
Private Sub Class_Initialize()
   m_Length = 0
   m_BlockSize = DEFAULT_BLOCK_SIZE
   m_Size = 0
End Sub
