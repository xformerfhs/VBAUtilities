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
'| Description         | Convert byte array from and to Base64 strings
'|---------------------+---------------------------------------------------
'| Author              | Frank Schwab
'|---------------------+---------------------------------------------------
'| Version             | 1.0.0
'|---------------------+---------------------------------------------------
'| Changes             | 2020-09-10  Created. fhs
'|---------------------+---------------------------------------------------
'| Remarks             | ./.
'+-------------------------------------------------------------------------

Option Explicit

'
' Constants for error message
'
Private Const ERR_STR_CLASS_NAME As String = "Base64Converter"

Private Const ERR_NUM_BASE As Long = vbObjectError + 1842

Private Const ERR_NUM_API_ERROR As Long = ERR_NUM_BASE
Private Const ERR_STR_API_ERROR As String = "Unable to convert. '%1' returned code 0x%2: %3"

Private Const ERR_NUM_INVALID_LENGTH As Long = ERR_NUM_BASE + 1
Private Const ERR_STR_INVALID_LENGTH As String = "Base64 string has invalid length"

'
' Private constants
'
Private Const BASE64_PAD_CHARACTER As String = "="

'
' Windows API constants
'
Private Const CRYPT_STRING_BASE64 As Long = 1
Private Const CRYPT_STRING_NOCRLF As Long = &H40000000

'
' Windows API declarations
'
Private Declare PtrSafe Function CryptBinaryToString Lib "Crypt32" Alias "CryptBinaryToStringW" ( _
     ByVal pbBinary As LongPtr, _
     ByVal cbBinary As Long, _
     ByVal dwFlags As Long, _
     ByVal pszStringPtr As LongPtr, _
     ByRef pcchString As LongPtr) As Long

Private Declare PtrSafe Function CryptStringToBinary Lib "Crypt32.dll" Alias "CryptStringToBinaryW" ( _
     ByVal pszStringPtr As LongPtr, _
     ByVal cchString As Long, _
     ByVal dwFlags As Long, _
     ByVal pbBinary As LongPtr, _
     ByRef pcbBinary As LongPtr, _
     ByVal pdwSkip As LongPtr, _
     ByVal pdwFlags As LongPtr) As Long

'
' Instance variables
'
Private m_MM As New MessageManager

'
' Private methods
'

'
'+--------------------------------------------------------------------------
'| Method           | HandleAPIError
'|------------------+-------------------------------------------------------
'| Description      | Handle Windows API errors that set Err.LastDLLError.
'|------------------+-------------------------------------------------------
'| Parameters       | className  : Class the calling method belongs to
'|                  | errorNumber: The error number that should be raised
'|                  | functionName : The name of the API function that failed
'|------------------+-------------------------------------------------------
'| Return values    | ./.
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2020-09-09  Created. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | This method does not return but raises an error
'|                  | with the value of "Err.LastDLLError" as the API
'|                  | error code.
'+--------------------------------------------------------------------------
'
Private Sub HandleAPIError(ByRef apiFunctionName As String)
   Dim errorCode As Long
   errorCode = Err.LastDllError

   Err.Raise ERR_NUM_API_ERROR, _
             ERR_STR_CLASS_NAME, _
             m_MM.FormatMessageWithParameters(ERR_STR_API_ERROR, _
                                              apiFunctionName, _
                                              Hex$(errorCode), _
                                              m_MM.GetMessageForWindowsErrorCode(errorCode))
End Sub

'
'+--------------------------------------------------------------------------
'| Method           | StripBase64Padding
'|------------------+-------------------------------------------------------
'| Description      | Remove trailing padding characters from Base64 string
'|------------------+-------------------------------------------------------
'| Parameters       | base64TextWithPadding: A Base64 encoded string with
'|                  |                        padding characters
'|------------------+-------------------------------------------------------
'| Return values    | Base64 encoded string without padding characters
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2020-09-10  Created. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | ./.
'+--------------------------------------------------------------------------
'
Private Function StripBase64Padding(ByRef base64TextWithPadding As String) As String
   Dim textLength As Long

   textLength = Len(base64TextWithPadding)

   If textLength > 0 Then
      If (textLength And 3) = 0 Then
         Dim pos As Long
         Dim i As Long

         ' Here textLength has at least the value 4
         For i = textLength To textLength - 2 Step -1
            pos = i

            If Mid$(base64TextWithPadding, i, 1) <> BASE64_PAD_CHARACTER Then _
               Exit For
         Next

        If pos = textLength Then
           StripBase64Padding = base64TextWithPadding
        Else
           StripBase64Padding = Left$(base64TextWithPadding, pos)
        End If
      End If
   Else
      StripBase64Padding = base64TextWithPadding
   End If
End Function

'
'+--------------------------------------------------------------------------
'| Method           | AddBase64Padding
'|------------------+-------------------------------------------------------
'| Description      | Add trailing padding characters to Base64 string
'|------------------+-------------------------------------------------------
'| Parameters       | base64TextWithoutPadding: A Base64 encoded string
'|                  |                           without padding characters
'|------------------+-------------------------------------------------------
'| Return values    | Base64 encoded string with padding characters
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2020-09-10  Created. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | Right now this method is not used as the Windows API
'|                  | function CryptStringToBinary is able to cope with
'|                  | unpadded Base64 strings.
'+--------------------------------------------------------------------------
'
Private Function AddBase64Padding(ByRef base64TextWithoutPadding As String) As String
   Dim textLength As Long

   textLength = Len(base64TextWithoutPadding)

   If textLength > 0 Then
      textLength = textLength And 3

      If textLength <> 0 Then
         AddBase64Padding = base64TextWithoutPadding & String$(4 - textLength, BASE64_PAD_CHARACTER)
      Else
         AddBase64Padding = base64TextWithoutPadding
      End If
   Else
      AddBase64Padding = base64TextWithoutPadding
   End If
End Function

'
' Public methods
'

'
'+--------------------------------------------------------------------------
'| Method           | Encode
'|------------------+-------------------------------------------------------
'| Description      | Encode a byte array as a Base64 encoded string
'|------------------+-------------------------------------------------------
'| Parameters       | sourceByteArray: Byte array to encode
'|------------------+-------------------------------------------------------
'| Return values    | Base64 encoded byte array string
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2020-09-08  Created. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | ./.
'+------------------+-------------------------------------------------------
'| Typical usage    | Dim b64 As New Base64Converter
'|                  | Dim aSourceArray(1 to 100) As Byte
'|                  | Dim b64Coded As String
'|                  | ...
'|                  | b64Coded = b64.Encode(aSourceArray)
'+--------------------------------------------------------------------------
'
Public Function Encode(ByRef sourceByteArray() As Byte) As String
   ' Calculate length of source byte array
   Dim sourceLength As Long
   sourceLength = GetArrayLength(sourceByteArray)

   Dim resultB64Text As String
   Dim b64TextLength As Long

   If sourceLength > 0 Then
      ' Get pointer to source byte array
      Dim sourcePointer As Long
      sourcePointer = VarPtr(sourceByteArray(LBound(sourceByteArray)))

      ' Calculate the size of the resulting Base64 string including the terminating null character
      b64TextLength = ((((sourceLength - 1) \ 3) + 1) * 4) + 1

      ' Initialize receiving string
      resultB64Text = Space$(b64TextLength)

      ' Get pointer to receiving string
      Dim b64TextPointer As Long
      b64TextPointer = StrPtr(resultB64Text)

      '
      ' Call the Windows API function to convert the source byte array into a base64 encoded string
      '
      Dim rc As Long

      rc = CryptBinaryToString(sourcePointer, _
                               sourceLength, _
                               CRYPT_STRING_BASE64 Or CRYPT_STRING_NOCRLF, _
                               b64TextPointer, _
                               b64TextLength)

      ' Check for error and raise an exception if there was one
      If rc = BOOL_API_ERROR Then _
         HandleAPIError "CryptBinaryToString"
   End If

   Encode = Left$(resultB64Text, b64TextLength)
End Function

'
'+--------------------------------------------------------------------------
'| Method           | EncodeNoPadding
'|------------------+-------------------------------------------------------
'| Description      | Encode a byte array as a Base64 encoded string
'|                  | without trailing padding characters
'|------------------+-------------------------------------------------------
'| Parameters       | sourceByteArray: Byte array to encode
'|------------------+-------------------------------------------------------
'| Return values    | Base64 encoded byte array string without padding
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2020-09-10  Created. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | ./.
'+------------------+-------------------------------------------------------
'| Typical usage    | Dim b64 As New Base64Converter
'|                  | Dim aSourceArray(1 To 100) As Byte
'|                  | Dim b64Coded As String
'|                  | ...
'|                  | b64Coded = b64.EncodeNoPadding(aSourceArray)
'+--------------------------------------------------------------------------
'
Public Function EncodeNoPadding(ByRef sourceByteArray() As Byte) As String
   EncodeNoPadding = StripBase64Padding(Encode(sourceByteArray))
End Function
'
'+--------------------------------------------------------------------------
'| Method           | Decode
'|------------------+-------------------------------------------------------
'| Description      | Decode a Base64 string into a byte array
'|------------------+-------------------------------------------------------
'| Parameters       | sourceB64Text: The Base64 encoded string to decode
'|------------------+-------------------------------------------------------
'| Return values    | Decoded byte array
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2020-09-08  Created. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | This method can decode padded and unpadded Base64
'|                  | string as CryptStringToBinary is able to cope
'|                  | with both.
'+------------------+-------------------------------------------------------
'| Typical usage    | Dim b64 As New Base64Converter
'|                  | Dim b64Coded As String
'|                  | Dim decodedArray() As Byte
'|                  | ...
'|                  | decodedArray = b64.Decode(b64Coded)
'+--------------------------------------------------------------------------
'
Public Function Decode(ByRef sourceB64Text As String) As Byte()
   ' Get source length
   Dim sourceLength As Long
   sourceLength = Len(sourceB64Text)

   Dim resultByteArray() As Byte

   If sourceLength > 0 Then
      ' The source length must never be a multiple of 4 plus 1 (i.e. 1, 5, 9, etc.)
      If (sourceLength And 3) <> 1 Then
         ' Get pointer to Base64 encoded string
         Dim b64TextPointer As Long
         b64TextPointer = StrPtr(sourceB64Text)

         ' Get (and overestimate when padding is used) the length of resulting byte array
         Dim resultByteArrayLength As Long
         resultByteArrayLength = (Len(sourceB64Text) * 3) \ 4  ' Allocate a bit too much, so we do not need to call CryptStringToBinary twice

         ' Allocate byte array
         ReDim resultByteArray(1 To resultByteArrayLength)

         ' Get pointer to result byte array
         Dim resultPointer As Long
         resultPointer = VarPtr(resultByteArray(1))

         Dim rc As Long

         rc = CryptStringToBinary(b64TextPointer, _
                                  Len(sourceB64Text), _
                                  CRYPT_STRING_BASE64, _
                                  resultPointer, _
                                  resultByteArrayLength, _
                                  0&, _
                                  0&)

         ' Check for error and raise an exception if there was one
         If rc = BOOL_API_ERROR Then _
            HandleAPIError "CryptStringToBinary"

         ' Set the size of the result byte array to the one returned by the CryptStringToBinary call
         ReDim Preserve resultByteArray(1 To resultByteArrayLength)
      Else
         Err.Raise ERR_NUM_INVALID_LENGTH, _
                   ERR_STR_CLASS_NAME, _
                   ERR_STR_INVALID_LENGTH
      End If
      
   End If

   Decode = resultByteArray
End Function

