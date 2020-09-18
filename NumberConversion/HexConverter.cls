VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "HexConverter"
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
'| Class               | HexConverter
'|---------------------+---------------------------------------------------
'| Description         | Convert byte arrays to and from hex strings
'|---------------------+---------------------------------------------------
'| Author              | Frank Schwab
'|---------------------+---------------------------------------------------
'| Version             | 1.1.0
'|---------------------+---------------------------------------------------
'| Changes             | 2014-04-02  Created. fhs
'|                     | 2014-07-18  Handle invalid hex char. fhs
'|---------------------+---------------------------------------------------
'| Typical usage       | Dim aByteArray(1 To 100) As Byte
'|                     | Dim aHC As New HexConverter
'|                     | ...
'|                     | Debug.Print aHC.Encode(aByteArray)
'+-------------------------------------------------------------------------

Option Compare Binary
Option Explicit

'
' Constants for error messages
'
Private Const STR_ERROR_SOURCE As String = "HexConverter"
Private Const ERR_BASE         As Long = vbObjectError + 21450

Private Const ERR_INVALID_HEX_STRING_LENGTH     As Long = ERR_BASE + 0
Private Const STR_ERR_INVALID_HEX_STRING_LENGTH As String = "Hexadecimal string has an odd length: "

Private Const ERR_EMPTY_STRING     As Long = ERR_BASE + 1
Private Const STR_ERR_EMPTY_STRING As String = "Hexadecimal string is empty"

Private Const ERR_INVALID_HEX_CHAR     As Long = ERR_BASE + 2
Private Const STR_ERR_INVALID_HEX_CHAR As String = "Invalid hexadecimal character(s): "


'
' Public methods
'

'
'+--------------------------------------------------------------------------
'| Method           | Encode
'|------------------+-------------------------------------------------------
'| Description      | Convert byte array to hex string
'|------------------+-------------------------------------------------------
'| Parameter        | ./.
'|------------------+-------------------------------------------------------
'| Return values    | ./.
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2014-04-02  Created. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | ./.
'+--------------------------------------------------------------------------
'
Public Function Encode(ByRef anArray() As Byte) As String
   Dim i As Long
   Dim result As String
   Dim size As Long
   Dim aValue As Byte
   Dim resultIndex As Long

   size = UBound(anArray) - LBound(anArray) + 1
   result = String$(size + size, "0")

   resultIndex = 1
   For i = LBound(anArray) To UBound(anArray)
      aValue = anArray(i)

      If aValue < &H10 Then
         resultIndex = resultIndex + 1
         Mid$(result, resultIndex, 1) = Hex$(aValue)
      Else
         Mid$(result, resultIndex, 2) = Hex$(aValue)
         resultIndex = resultIndex + 1
      End If

      resultIndex = resultIndex + 1
   Next i

   Encode = result
End Function

'
'+--------------------------------------------------------------------------
'| Method           | Decode
'|------------------+-------------------------------------------------------
'| Description      | Convert hex string to byte array
'|------------------+-------------------------------------------------------
'| Parameter        | ./.
'|------------------+-------------------------------------------------------
'| Return values    | ./.
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2014-04-02  Created. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | ./.
'+--------------------------------------------------------------------------
'
Public Function Decode(ByRef aHexString As String) As Byte()
   Dim size As Long
   Dim result() As Byte
   Dim resultIndex As Long
   Dim i As Long
   Dim textToConvert As String

   size = Len(aHexString)
   
   If size > 0 Then
      If (size And 1) = 0 Then
         ReDim result(1 To (size \ 2))

         resultIndex = LBound(result)

         On Error GoTo InvalidCharacterError

         For i = 1 To size Step 2
            '
            ' Converting from string to byte by prepending "&H" to the
            ' next two characters is more than 3 times faster than looking
            ' up each character via InStr and calculating the byte value
            ' from two such lookups.
            '
            textToConvert = Mid$(aHexString, i, 2)
            result(resultIndex) = CByte("&H" & textToConvert)
            resultIndex = resultIndex + 1
         Next i
      Else
         Err.Raise ERR_INVALID_HEX_STRING_LENGTH, _
                   STR_ERROR_SOURCE, _
                   STR_ERR_INVALID_HEX_STRING_LENGTH & Format$(size)
      End If
   Else
      Err.Raise ERR_EMPTY_STRING, _
                STR_ERROR_SOURCE, _
                STR_ERR_EMPTY_STRING
   End If

   Decode = result

   Exit Function

InvalidCharacterError:
   Err.Raise ERR_INVALID_HEX_CHAR, _
             STR_ERROR_SOURCE, _
             STR_ERR_INVALID_HEX_CHAR & textToConvert
End Function
