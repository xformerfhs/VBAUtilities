VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Base64Converter"
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
'| Class               | Base64Converter
'|---------------------+---------------------------------------------------
'| Description         | Base64 encoding and decoding
'|---------------------+---------------------------------------------------
'| Author              | Frank Schwab
'|---------------------+---------------------------------------------------
'| Version             | 3.2.0
'|---------------------+---------------------------------------------------
'| Changes             | 2006-02-25  Created. fhs
'|                     | 2006-03-01  V2: Added correct error handling. fhs
'|                     | 2010-05-07  V3: Converted to byte array. fhs
'|                     | 2014-03-19  General simplification. fhs
'|                     | 2014-07-24  Decoding simplified . fhs
'|                     | 2014-07-29  Added more error checking to decoding. fhs
'|---------------------+---------------------------------------------------
'| Remarks             | ./.
'|---------------------+---------------------------------------------------
'| Typical call        | Dim base64 As New Base64Converter
'|                     | Dim hashValue(1 To 20) As Byte
'|                     | ...
'|                     | ' Compute hash value
'|                     | ...
'|                     | Open "test.b64" For Output As #1
'|                     | Print #1, base64.Encode(hashValue)
'|                     | Close #1
'+-------------------------------------------------------------------------

Option Compare Binary
Option Explicit

'
' Private constants
'

'
' Constants for error messages
'
Private Const STR_ERR_SOURCE   As String = "Base64Converter"
Private Const ERR_BASE As Long = vbObjectError + 1842

' Invalid length for decoding
Private Const ERR_INVALID_LENGTH     As Long = ERR_BASE
Private Const STR_ERR_INVALID_LENGTH As String = "The length of the string to decode is not a multiple of 4: "

' Invalid character while decoding
Private Const ERR_INVALID_CHAR     As Long = ERR_BASE + 1
Private Const STR_ERR_INVALID_CHAR As String = "Invalid character at position "

'
' List of characters that will be put out for 6 bits: A=000000, ..., /=111111
'
Private Const ENCODING_CHARACTERS As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"

'
' Constants for padding character "=" if the length of the source bytes
' is not a multiple of 3
'
Private Const PADDING_CHARACTER       As String * 1 = "="
Private Const PADDING_CHARACTER_VALUE As Byte = &HFF

'
' Value for an invalid coding character
'
Private Const INVALID_CHARACTER_VALUE As Byte = &HA5


'
' Instance variables
'

'
' m_ValueToCharacter contains the mapping from a byte value to a character
'
Private m_ValueToCharacter(0 To (Len(ENCODING_CHARACTERS) - 1)) As String * 1

'
' m_CharacterToValue contains the mapping from a base64 character to
' the correspondig byte value. The boundaries 32...127 would be sufficient
' but then one would need complicated tests for invalid characters.
' Having a value for every character value from 0 to 255 makes
' error checking much easier.
'
Private m_CharacterToValue(0 To 255) As Byte

'
' Private methods
'

'
'+--------------------------------------------------------------------------
'| Method           | GetCharacterValue
'|------------------+-------------------------------------------------------
'| Decription       | Gets code of a base64 character
'|------------------+-------------------------------------------------------
'| Parameter        | aCharacter: Character to convert to code
'|------------------+-------------------------------------------------------
'| Return values    | Byte value of the character
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2014-07-29  Created. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | ./.
'+--------------------------------------------------------------------------
'
Private Function GetCharacterValue(ByRef aCharacter As String) As Integer
   Dim code As Integer

   code = Asc(aCharacter)

   If code < LBound(m_CharacterToValue) Then
      GetCharacterValue = INVALID_CHARACTER_VALUE
   Else
      If code > UBound(m_CharacterToValue) Then
         GetCharacterValue = INVALID_CHARACTER_VALUE
      Else
         GetCharacterValue = m_CharacterToValue(code)
      End If
   End If
End Function

'
' Public methods
'

'
'+--------------------------------------------------------------------------
'| Method           | Encode
'|------------------+-------------------------------------------------------
'| Decription       | Convert bytes to a base64 string
'|------------------+-------------------------------------------------------
'| Parameter        | inputByteArray: Array with bytes to convert
'|------------------+-------------------------------------------------------
'| Return values    | Base64 string
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2006-02-25  Created. fhs
'|                  | 2010-05-07  Use byte array as input. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | ./.
'+--------------------------------------------------------------------------
'
Public Function Encode(inputByteArray() As Byte) As String
   Dim result As String

   Dim value(1 To 3) As Byte

   Dim i1 As Byte, i2 As Byte, i3 As Byte, i4 As Byte

   Dim inputIndex As Long
   Dim outputIndex As Long
   Dim codeIndex As Long

   Dim validBytesCount As Integer
   Dim inputSize As Long

   Dim resultSize As Long

   result = ""

   inputSize = UBound(inputByteArray) - LBound(inputByteArray) + 1

   If inputSize > 0 Then
      '
      ' Calculate size of resulting string:
      ' resultSize = 4 * RoundUp(inputSize / 3), i.e. 3 input bytes generate 4 output bytes
      '
      resultSize = inputSize \ 3

      If (resultSize * 3) <> inputSize Then
         resultSize = resultSize + 1
      End If

      resultSize = resultSize * 4

      '
      ' Allocate result string. This is faster than always appending a new character, as appending
      ' involves a copy operation of the old string every time a character is appended.
      '
      result = String$(resultSize, "A")

      validBytesCount = UBound(value) - LBound(value) + 1

      inputIndex = LBound(inputByteArray)
      outputIndex = 1

      Do
         '
         ' Always process a block of three bytes.
         '
         ' If there are fewer bytes left then fill the block with zeroes.
         '
         For codeIndex = LBound(value) To UBound(value)
            If inputIndex <= inputSize Then
               value(codeIndex) = inputByteArray(inputIndex)
               inputIndex = inputIndex + 1
            Else
               value(codeIndex) = 0
               validBytesCount = validBytesCount - 1
            End If
         Next codeIndex

         '
         ' From the three bytes calculate four indexes into m_ValueToCharacter
         '

         '
         ' Doing the shifts with arithmetics is 9 times faster than looking up
         ' the shift values up in a precomputed shift table!
         '
         i1 = (value(1) \ 4) And &H3F                                 ' First 6 bits of first byte
         i2 = ((value(1) * 16) And &H30) Or ((value(2) \ 16) And &HF) ' Last 2 bits of first byte and first 4 bits of second byte
         i3 = ((value(2) * 4) And &H3C) Or ((value(3) \ 64) And &H3)  ' Last 4 bits of second byte and first 2 bits of third byte
         i4 = value(3) And &H3F                                       ' Last 6 bits of third byte

         '
         ' 3 bytes of input data always yield 4 characters.
         '
         ' 1 byte with 8 bits results in at least 2 output characters with 6 bits in the first
         ' and 2 bits in the second character.
         '
         Mid$(result, outputIndex, 1) = m_ValueToCharacter(i1)
         outputIndex = outputIndex + 1

         Mid$(result, outputIndex, 1) = m_ValueToCharacter(i2)
         outputIndex = outputIndex + 1

         '
         ' If the input block is just 1 byte 2 padding characters must be put out.
         '
         Mid$(result, outputIndex, 1) = IIf(validBytesCount > 1, m_ValueToCharacter(i3), PADDING_CHARACTER)
         outputIndex = outputIndex + 1

         '
         ' If the input block is 2 bytes 1 padding character must be put out.
         '
         Mid$(result, outputIndex, 1) = IIf(validBytesCount > 2, m_ValueToCharacter(i4), PADDING_CHARACTER)
         outputIndex = outputIndex + 1
      Loop Until inputIndex > inputSize
   End If
   
   Encode = result
End Function

'+--------------------------------------------------------------------------
'| Method           | Decode
'|------------------+-------------------------------------------------------
'| Decription       | Convert a base64 string to a byte array
'|------------------+-------------------------------------------------------
'| Parameter        | inputString: Base64 string
'|------------------+-------------------------------------------------------
'| Return values    | Decoded byte array
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2006-02-25  Created. fhs
'|                  | 2010-05-07  Use byte array as output. fhs
'|                  | 2014-07-24  Simplified. fhs
'|                  | 2014-07-29  Calculate resulting array size before
'|                  |             starting processing. Added checks for
'|                  |             padding characters at invalid positions.
'|                  |             Added exception when string is
'|                  |             too short. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | Decoding needs a lot of error checking so this
'|                  | method is the one that throws errors, at all.
'|                  | And it is much more complex and has more checks
'|                  | than the encoding method.
'+--------------------------------------------------------------------------
'
Public Function Decode(inputString As String) As Byte()
   Dim result() As Byte
   
   Dim n1 As Byte, n2 As Byte, n3 As Byte, n4 As Byte
   
   Dim inputSize As Long

   Dim currentPosition As Long
   Dim endPosition As Long
   Dim errorPosition As Long
   
   Dim resultIndex As Long

   Dim hasInvalidCharacter As Boolean
   Dim inLastBlock As Boolean

   inputSize = Len(inputString)
   
   If inputSize > 3 Then
      If (inputSize And 3) = 0 Then          ' "(inputSize And 3)" is only 0 if the length of the input string is a multiple of 4
         currentPosition = 1
         endPosition = inputSize - 3         ' We always process 4 characters at a time so this is the max. position for the last block

         resultIndex = (inputSize \ 4) * 3   ' Calculate maximum size of the result array

         '
         ' Correct for one or two padding characters. In a valid base64 string there
         ' can be no more padding characters.
         '
         If Right$(inputString, 1) = PADDING_CHARACTER Then ' Correct for one padding character
            resultIndex = resultIndex - 1
         End If

         If Mid$(inputString, inputSize - 1, 1) = PADDING_CHARACTER Then ' Correct for two padding character
            resultIndex = resultIndex - 1
         End If

         ReDim result(1 To resultIndex)   ' Allocate resulting byte array

         resultIndex = 0
   
         errorPosition = 0

         hasInvalidCharacter = False
         inLastBlock = False

         Do
            '
            ' Always take a block of 4 characters and convert them to 3 bytes.
            '
            ' First we get the 4 numbers that the characters correspond to.
            ' Each number is a 6 bit value. The most significant 2 bits are always 0.
            '
            ' We need to check for invalid characters at each position and the
            ' first two characters of each block must not be padding characters.
            '
            n1 = GetCharacterValue(Mid$(inputString, currentPosition, 1))

            If n1 <> INVALID_CHARACTER_VALUE Then
               If n1 <> PADDING_CHARACTER_VALUE Then
                  currentPosition = currentPosition + 1

                  n2 = GetCharacterValue(Mid$(inputString, currentPosition, 1))

                  If n2 <> INVALID_CHARACTER_VALUE Then
                     If n2 <> PADDING_CHARACTER_VALUE Then
                        currentPosition = currentPosition + 1

                        n3 = GetCharacterValue(Mid$(inputString, currentPosition, 1))

                        If n3 <> INVALID_CHARACTER_VALUE Then
                           currentPosition = currentPosition + 1

                           n4 = GetCharacterValue(Mid$(inputString, currentPosition, 1))

                           If n4 <> INVALID_CHARACTER_VALUE Then
                              currentPosition = currentPosition + 1
                           Else
                              errorPosition = currentPosition
                           End If
                        Else
                           errorPosition = currentPosition
                        End If
                     Else
                        errorPosition = currentPosition
                     End If
                  Else
                     errorPosition = currentPosition
                  End If
               Else
                  errorPosition = currentPosition
               End If
            Else
               errorPosition = currentPosition
            End If

            '
            ' Set check variables
            '
            hasInvalidCharacter = (errorPosition <> 0)
            inLastBlock = (currentPosition > endPosition)

            '
            ' If there was no error process the 4 bytes
            '
            If Not hasInvalidCharacter Then
               '
               ' The first byte is built from the first 2 characters
               '
               resultIndex = resultIndex + 1
               result(resultIndex) = (n1 * 4) Or (n2 \ 16)  ' 6 bits from first value and 2 bits from second value
   
               '
               ' The next bytes are only converted if they are no padding bytes
               '
               If n3 <> PADDING_CHARACTER_VALUE Then
                  resultIndex = resultIndex + 1
                  result(resultIndex) = ((n2 And &HF) * 16) Or (n3 \ 4)  ' 4 bits from second value and 4 bits from third value
               Else
                  If Not inLastBlock Then
                     errorPosition = currentPosition - 2    ' There must be no padding characters in a block that is not the last one
                     hasInvalidCharacter = True
                  End If
               End If ' n3 <> PADDING_CHARACTER_VALUE
   
               If n4 <> PADDING_CHARACTER_VALUE Then
                  resultIndex = resultIndex + 1
                  result(resultIndex) = ((n3 And &H3) * 64) Or n4  ' 2 bits from third value and 6 bits from fourth value
               Else
                  If Not inLastBlock Then
                     errorPosition = currentPosition - 1    ' There must be no padding characters in a block that is not the last one
                     hasInvalidCharacter = True
                  End If
               End If ' n4 <> PADDING_CHARACTER_VALUE
            End If    ' errorPosition = 0
         Loop Until inLastBlock Or hasInvalidCharacter

         '
         ' If there was an invalid character, throw an exception
         '
         If hasInvalidCharacter Then
            Err.Raise ERR_INVALID_CHAR, _
                      STR_ERR_SOURCE, _
                      STR_ERR_INVALID_CHAR & _
                         Format$(errorPosition) & _
                         ": '" & _
                         Mid$(inputString, errorPosition, 1) & _
                         "' in string '" & _
                         inputString & _
                         "'"

         End If
      Else ' (inputSize And 3) = 0
         '
         ' Input length was not a multiple of 4, so throw an exception
         '
         Err.Raise ERR_INVALID_LENGTH, _
                   STR_ERR_SOURCE, _
                   STR_ERR_INVALID_LENGTH & _
                      "'" & _
                      inputString & _
                      "'"
      End If ' (inputSize And 3) = 0
   Else
      '
      ' Input length was less than 4, so throw an exception
      '
      Err.Raise ERR_INVALID_LENGTH, _
                STR_ERR_SOURCE, _
                STR_ERR_INVALID_LENGTH & _
                   "'" & _
                   inputString & _
                   "'"
   End If ' inputSize > 3

   Decode = result
End Function


'
' Class methods
'

'
'+--------------------------------------------------------------------------
'| Method           | Class_Initialize
'|------------------+-------------------------------------------------------
'| Decription       | Constructor
'|------------------+-------------------------------------------------------
'| Parameter        | ./.
'|------------------+-------------------------------------------------------
'| Return values    | ./.
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2002-03-15  Created. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | Initialize encoding and decoding arrays.
'+--------------------------------------------------------------------------

Private Sub Class_Initialize()
   Dim i As Long
   Dim ch As String * 1

   For i = LBound(m_CharacterToValue) To UBound(m_CharacterToValue)
      m_CharacterToValue(i) = INVALID_CHARACTER_VALUE
   Next i
   
   For i = 0 To Len(ENCODING_CHARACTERS) - 1
      ch = Mid(ENCODING_CHARACTERS, i + 1, 1)
      
      m_ValueToCharacter(i) = ch
      m_CharacterToValue(Asc(ch)) = i
   Next i
   
   m_CharacterToValue(Asc(PADDING_CHARACTER)) = PADDING_CHARACTER_VALUE
End Sub
