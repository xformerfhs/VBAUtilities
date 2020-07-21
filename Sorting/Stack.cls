VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Stack"
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
'| Class               | Stack
'|---------------------+---------------------------------------------------
'| Description         | An implementation of a stack
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
' Private Konstanten für Fehlermeldungen
'
Private Const MODULNAME As String = "Stack"

Private Const N_START_ERROR_MESSAGE As Long = vbObjectError + 2600

Private Const ERR_STACK_IS_EMPTY     As Long = N_START_ERROR_MESSAGE + 0
Private Const STR_ERR_STACK_IS_EMPTY As String = "Stack is empty"

'
' Private Konstanten
'
Private Const DEFAULT_BLOCKSIZE As Long = 50
Private Const MIN_BLOCKSIZE     As Long = 10
Private Const MAX_BLOCKSIZE     As Long = 100000

'
' Instanzvariable
'
Private m_Items() As Variant
Private m_BlockSize As Long
Private m_ActSize As Long
Private m_ActIndex As Long

'
' Private Methoden
'

'
' Öffentliche Methoden
'
Public Property Get BlockSize() As Long
   BlockSize = m_BlockSize
End Property

Public Property Let BlockSize(ByVal newBlockSize As Long)
   If newBlockSize < MIN_BLOCKSIZE Then _
      newBlockSize = MIN_BLOCKSIZE

   If newBlockSize > MAX_BLOCKSIZE Then _
      newBlockSize = MAX_BLOCKSIZE

   m_BlockSize = newBlockSize
End Property

Public Property Get IsEmpty() As Boolean
   IsEmpty = (m_ActIndex = 0)
End Property

Public Property Get Depth() As Long
   Depth = m_ActIndex
End Property

Public Property Get TopElement() As Variant
   If Me.IsEmpty Then
      Err.Raise ERR_STACK_IS_EMPTY, MODULNAME, STR_ERR_STACK_IS_EMPTY
   Else
      TopElement = m_Items(m_ActIndex)
   End If
End Property

Public Sub Reset()
   m_ActIndex = 0
End Sub

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

Public Function Pop() As Variant
   If Me.IsEmpty Then
      Err.Raise ERR_STACK_IS_EMPTY, MODULNAME, STR_ERR_STACK_IS_EMPTY
   Else
      Pop = m_Items(m_ActIndex)
      m_ActIndex = m_ActIndex - 1
   End If
End Function

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
