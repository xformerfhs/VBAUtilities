VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Stack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False'
'+-------------------------------------------------------------------------
'|
'| SPDX-FileCopyrightText: 2020 Frank Schwab
'|
'| SPDX-License-Identifier: MIT
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
'| Class               | Stack
'|---------------------+---------------------------------------------------
'| Description         | An implementation of a stack
'|---------------------+---------------------------------------------------
'| Author              | Frank Schwab
'|---------------------+---------------------------------------------------
'| Version             | 1.1.0
'|---------------------+---------------------------------------------------
'| Changes             | 2018-09-25  Created. fhs
'|                     | 2020-09-20  Refactored. fhs
'|---------------------+---------------------------------------------------
'| Remarks             | This implementation allocates space for the
'|                     | stack in blocks, the size of which can be set.
'+-------------------------------------------------------------------------
'

Option Explicit

'
' Constants for error messages
'
Private Const ERR_STR_CLASS_NAME As String = "Stack"

Private Const ERR_NUM_START As Long = vbObjectError + 2600

Private Const ERR_NUM_STACK_IS_EMPTY As Long = ERR_NUM_START
Private Const ERR_STR_STACK_IS_EMPTY As String = "Stack is empty"

Private Const ERR_NUM_INVALID_BLOCK_SIZE As Long = ERR_NUM_START + 1
Private Const ERR_STR_INVALID_BLOCK_SIZE As String = "Invalid block size"

'
' Private constants
'
Private Const DEFAULT_BLOCKSIZE As Long = 50
Private Const MIN_BLOCKSIZE     As Long = 10
Private Const MAX_BLOCKSIZE     As Long = 100000

'
' Instance variables
'
Private m_Items() As Variant
Private m_BlockSize As Long
Private m_ActSize As Long
Private m_ActIndex As Long


'
' Public attributes
'

'
'+--------------------------------------------------------------------------
'| Method           | BlockSize
'|------------------+-------------------------------------------------------
'| Description      | Get block size of stack
'|------------------+-------------------------------------------------------
'| Parameter        | ./.
'|------------------+-------------------------------------------------------
'| Return values    | Block size of stack
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2018-09-25  Created. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | ./.
'+--------------------------------------------------------------------------
'
Public Property Get BlockSize() As Long
   BlockSize = m_BlockSize
End Property

'
'+--------------------------------------------------------------------------
'| Method           | BlockSize
'|------------------+-------------------------------------------------------
'| Description      | Set new block size for stack
'|------------------+-------------------------------------------------------
'| Parameter        | newBlockSize: New block size for stack
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
Public Property Let BlockSize(ByVal newBlockSize As Long)
   If (newBlockSize < MIN_BLOCKSIZE) Or _
      (newBlockSize > MAX_BLOCKSIZE) Then _
      Err.Raise ERR_NUM_INVALID_BLOCK_SIZE, _
                ERR_STR_CLASS_NAME, _
                ERR_STR_INVALID_BLOCK_SIZE

   m_BlockSize = newBlockSize
End Property

'
'+--------------------------------------------------------------------------
'| Method           | IsEmpty
'|------------------+-------------------------------------------------------
'| Description      | Check if stack is empty
'|------------------+-------------------------------------------------------
'| Parameter        | ./.
'|------------------+-------------------------------------------------------
'| Return values    | True : Stack is empty
'|                  | False: Stack is not empty
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2018-09-25  Created. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | ./.
'+--------------------------------------------------------------------------
'
Public Property Get IsEmpty() As Boolean
   IsEmpty = (m_ActIndex = 0)
End Property

'
'+--------------------------------------------------------------------------
'| Method           | Depth
'|------------------+-------------------------------------------------------
'| Description      | Get stack depth
'|------------------+-------------------------------------------------------
'| Parameter        | ./.
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
Public Property Get Depth() As Long
   Depth = m_ActIndex
End Property

'
'+--------------------------------------------------------------------------
'| Method           | Peek
'|------------------+-------------------------------------------------------
'| Description      | Get top element from stack (without popping it)
'|------------------+-------------------------------------------------------
'| Parameter        | ./.
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
Public Function Peek() As Variant
   If Me.IsEmpty Then
      Err.Raise ERR_NUM_STACK_IS_EMPTY, ERR_STR_CLASS_NAME, ERR_STR_STACK_IS_EMPTY
   Else
      Peek = m_Items(m_ActIndex)
   End If
End Function

'
'+--------------------------------------------------------------------------
'| Method           | Reset
'|------------------+-------------------------------------------------------
'| Description      | Reset (and empty) stack
'|------------------+-------------------------------------------------------
'| Parameter        | ./.
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
Public Sub Reset()
   m_ActIndex = 0
End Sub

'
'+--------------------------------------------------------------------------
'| Method           | Push
'|------------------+-------------------------------------------------------
'| Description      | Push value on stack
'|------------------+-------------------------------------------------------
'| Parameter        | var: Value to push
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
Public Sub Push(ByVal var As Variant)
   Dim newIndex As Long

   newIndex = m_ActIndex + 1

   If newIndex > m_ActSize Then
      m_ActSize = m_ActSize + m_BlockSize
      ReDim Preserve m_Items(1 To m_ActSize)
   End If

   m_Items(newIndex) = var

   m_ActIndex = newIndex
End Sub


'
'+--------------------------------------------------------------------------
'| Method           | Pop
'|------------------+-------------------------------------------------------
'| Description      | Pop value from stack
'|------------------+-------------------------------------------------------
'| Parameter        | ./.
'|------------------+-------------------------------------------------------
'| Return values    | Topmost value of stack
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2018-09-25  Created. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | ./.
'+--------------------------------------------------------------------------
'
Public Function Pop() As Variant
   If Me.IsEmpty Then
      Err.Raise ERR_NUM_STACK_IS_EMPTY, ERR_STR_CLASS_NAME, ERR_STR_STACK_IS_EMPTY
   Else
      Pop = m_Items(m_ActIndex)
      m_ActIndex = m_ActIndex - 1
   End If
End Function

'
'+--------------------------------------------------------------------------
'| Method           | Dump
'|------------------+-------------------------------------------------------
'| Description      | Get string with stack content
'|------------------+-------------------------------------------------------
'| Parameter        | delim: Optional delimiter
'|------------------+-------------------------------------------------------
'| Return values    | String with stack content
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2018-09-25  Created. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | ./.
'+--------------------------------------------------------------------------
'
Public Function Dump(Optional ByVal delim As String = ", ") As String
   Dim dumpText As String
   Dim i As Long

   If Me.IsEmpty Then
      dumpText = "<Empty>"
   Else
      For i = 1 To m_ActIndex
         dumpText = dumpText & delim & Format$(m_Items(i))
      Next i
   End If

   Dump = Right$(dumpText, Len(dumpText) - Len(delim))
End Function

'
' Class methods
'
Private Sub Class_Initialize()
   m_BlockSize = DEFAULT_BLOCKSIZE
   m_ActSize = 0
   m_ActIndex = 0
End Sub
