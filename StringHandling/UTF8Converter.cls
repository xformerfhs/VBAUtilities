VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UTF8Converter"
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
'| Class               | UTF8Converter
'|---------------------+---------------------------------------------------
'| Description         | Convert a VBA string from and to UTF-8
'|---------------------+---------------------------------------------------
'| Author              | Frank Schwab
'|---------------------+---------------------------------------------------
'| Version             | 1.0.4
'|---------------------+---------------------------------------------------
'| Changes             | 2006-03-06  Created. fhs
'|                     | 2020-07-19  Simplified array check. fhs
'|                     | 2020-08-03  Correct array check because it was a hack. fhs
'|                     | 2020-09-01  Made 64 bit compatible. fhs
'|                     | 2020-09-09  Simplified. fhs
'|---------------------+---------------------------------------------------
'| Remarks             | VBA stores strings in the UTF-16 encoding, i.e. each
'|                     | character uses two bytes. However, for calling external
'|                     | DLLs VBA always converts these strings to the Windows
'|                     | locale (ANSI) where each character occupies one byte.
'|                     |
'|                     | If one would write "ByVal lpWideCharStr As String"
'|                     | in the API declarations VBA would first convert this
'|                     | string to ANSI and then pass it on to the DLL function.
'|                     | In exactly the same way each returned string would first be
'|                     | converted from ANSI to UTF-16 and then this UTF-16
'|                     | string would be passed to the calling VBA function.
'|                     |
'|                     | The Windows API functions that are used in this class
'|                     | expect UTF-16 strings. If the VBA UTF-16-ANSI conversion
'|                     | would take place one would get all kinds of errors and
'|                     | even VBA crashes. So VBA must be prevented from converting
'|                     | between the different encodings.
'|                     |
'|                     | This is the reason why the pointers to the VBA strings
'|                     | are declared as "ByVal lpWideCharStr As Long". The correct
'|                     | pointer is determined by the undocumented VBA function
'|                     | "StrPtr" and then passed by value to the Windows API
'|                     | function. This way VBA can not interfere with it's
'|                     | UTF-16-ANSI conversion.
'|---------------------+---------------------------------------------------
'| Typical usage       | Dim u8c As New UTF8Converter
'|                     | Dim utf8Text() As Byte
'|                     | utf8Text = u8c.FromVBToUTF8("This is a text")
'|                     | ...
'|                     | Dim aText As String
'|                     | aText = u8c.FromUTF8ToVB(utf8Text)
'+-------------------------------------------------------------------------

Option Explicit


'
' Constants for error messages
'
Private Const ERR_STR_CLASS_NAME   As String = "UTF8Converter"
Private Const ERR_NUM_BASE         As Long = vbObjectError + 19907

Private Const ERR_NUM_INVALID_UTF16_CODE As Long = ERR_NUM_BASE
Private Const ERR_STR_INVALID_UTF16_CODE As String = "A VB character is not a valid UTF-16 code point"

Private Const ERR_NUM_INVALID_UTF8_CODE As Long = ERR_NUM_BASE + 1
Private Const ERR_STR_INVALID_UTF8_CODE As String = "Invalid UTF-8 byte sequence"

Private Const ERR_NUM_API As Long = ERR_NUM_BASE + 2
Private Const ERR_STR_API As String = "Unable to convert. '%1' returned code 0x%2: %3"


'
' API declarations
'
Private Declare PtrSafe Function WideCharToMultiByte Lib "kernel32.dll" ( _
                         ByVal CodePage As Long, _
                         ByVal dwFlags As Long, _
                         ByVal lpWideCharStr As LongPtr, _
                         ByVal cchWideChar As Long, _
                         ByRef lpMultiByteStr As Byte, _
                         ByVal cbMultiByte As Long, _
                         ByVal lpDefaultChar As LongPtr, _
                         ByVal lpUsedDefaultChar As LongPtr) As Long

Private Declare PtrSafe Function MultiByteToWideChar Lib "kernel32.dll" ( _
                         ByVal CodePage As Long, _
                         ByVal dwFlags As Long, _
                         ByRef lpMultiByteStr As Byte, _
                         ByVal cbMultiByte As Long, _
                         ByVal lpWideCharStr As LongPtr, _
                         ByVal cchWideChar As Long) As Long

                        
'
' API constants
'
Private Const CP_UTF8 As Long = 65001

Private Const WC_ERR_INVALID_CHARS As Long = &H80

Private Const MB_ERR_INVALID_CHARS As Long = &H8

Private Const ERROR_NO_UNICODE_TRANSLATION As Long = 1113

'
' Instance variables
'
Private m_MM As New MessageManager

'
' Private methods
'

'
'+--------------------------------------------------------------------------
'| Method           | GetArrayLength
'|------------------+-------------------------------------------------------
'| Description      | Get length of a byte array
'|------------------+-------------------------------------------------------
'| Parameter        | aByteArray: Byte array to get the length for
'|------------------+-------------------------------------------------------
'| Return values    | Length of byte array
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2020-07-18  Created. fhs
'|                  | 2020-08-03  Correct array check because it was a hack. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | When one tries to calculate the length of an
'|                  | uninitialized array an error is raised.
'|                  | This function catches this error and returns a
'|                  | length of 0, instead.
'+--------------------------------------------------------------------------
'
Private Function GetArrayLength(ByRef aByteArray() As Byte) As Long
   On Error Resume Next

   GetArrayLength = UBound(aByteArray) - LBound(aByteArray) + 1

   ' If the array is empty there was an error so we set
   ' the return value accordingly

   If Err.Number <> 0 Then _
      GetArrayLength = 0
End Function

'
'+--------------------------------------------------------------------------
'| Method           | HandleAPIError
'|------------------+-------------------------------------------------------
'| Description      | Raise an error from for Windows API calls that set
'|                  | Err.LastDLLError
'|------------------+-------------------------------------------------------
'| Parameters       | apiFunctionName: The name of the function that returned
'|                  |                  an error.
'|------------------+-------------------------------------------------------
'| Return values    | None.
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2006-03-06  Created. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | This method does not return because it raises an
'|                  | error.
'+--------------------------------------------------------------------------
'
Private Sub HandleAPIError(ByRef apiFunctionName As String)
   Dim errorCode As Long
   
   errorCode = Err.LastDllError
      
   Err.Raise ERR_NUM_API, _
             ERR_STR_CLASS_NAME, _
             m_MM.FormatMessageWithParameters(ERR_STR_API, _
                                              apiFunctionName, _
                                              Hex$(errorCode), _
                                              m_MM.GetMessageForWindowsErrorCode(errorCode))
End Sub

'
' Public methods
'

'
'+--------------------------------------------------------------------------
'| Method           | FromVBToUTF8
'|------------------+-------------------------------------------------------
'| Description      | Convert from UTF-16 to UTF-8
'|------------------+-------------------------------------------------------
'| Parameter        | aVBString: String to convert.
'|------------------+-------------------------------------------------------
'| Return values    | UTF-8 representation of UTF-16 string as byte array.
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2006-03-06  Created. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | ./.
'+--------------------------------------------------------------------------
'
Public Function FromVBToUTF8(ByRef aVBString As String) As Byte()
   Dim stringSize As Long
   Dim pointerToString As LongPtr
   Dim utf8Size As Long
   Dim result() As Byte

   stringSize = Len(aVBString)

   If stringSize > 0 Then
      '
      ' The pointer to the VBA string *must* be calculated by calling "StrPtr"
      ' otherwise VBA will convert to ANSI before calling the API function
      '
      pointerToString = StrPtr(aVBString)

      '
      ' First, get size of array to receive the UTF-8 representation
      '
      utf8Size = WideCharToMultiByte(CP_UTF8, _
                                     WC_ERR_INVALID_CHARS, _
                                     pointerToString, _
                                     stringSize, _
                                     ByVal 0&, _
                                     0&, _
                                     0&, _
                                     0&)

      If utf8Size > 0 Then
         '
         ' Set the size of the result array accordingly.
         '
         ' Getting it right the first time is 15% faster than allocating a
         ' too large array and correcting that after the conversion.
         '
         ReDim result(1 To utf8Size)

         '
         ' Now call the Windows API function again, but this time with
         ' the correctly sized array that receives the UTF-8 representation
         '
         WideCharToMultiByte CP_UTF8, _
                             0&, _
                             pointerToString, _
                             stringSize, _
                             result(LBound(result)), _
                             utf8Size, _
                             0&, _
                             0&
      Else
         If Err.LastDllError = ERROR_NO_UNICODE_TRANSLATION Then
            Err.Raise ERR_NUM_INVALID_UTF16_CODE, _
                      ERR_STR_CLASS_NAME, _
                      ERR_STR_INVALID_UTF16_CODE
         Else
            HandleAPIError "WideCharToMultiByte"
         End If
      End If
   End If

   FromVBToUTF8 = result
End Function

'
'+--------------------------------------------------------------------------
'| Method           | FromUTF8ToVB
'|------------------+-------------------------------------------------------
'| Description      | Convert from UTF-8 to UTF-16
'|------------------+-------------------------------------------------------
'| Parameter        | aVBString: String to convert.
'|------------------+-------------------------------------------------------
'| Return values    | UTF-8 representation of UTF-16 string as byte array.
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2006-03-06  Created. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | ./.
'+--------------------------------------------------------------------------
'
Public Function FromUTF8ToVB(ByRef utf8ByteArray() As Byte) As String
   Dim utf8Size As Long
   Dim resultSize As Long
   Dim result As String
   Dim rc As Long

   utf8Size = GetArrayLength(utf8ByteArray)
   
   If utf8Size > 0 Then
      '
      ' First, get size of the string to receive the UTF-16 representation
      '
      resultSize = MultiByteToWideChar(CP_UTF8, _
                                       MB_ERR_INVALID_CHARS, _
                                       utf8ByteArray(LBound(utf8ByteArray)), _
                                       utf8Size, _
                                       0&, _
                                       0&)

      If resultSize > 0 Then
         '
         ' Set the size of the result string accordingly
         '
         ' Getting it right the first time is 20% faster than allocating a
         ' too large string and correcting that after the conversion.
         '
         result = Space$(resultSize)

         '
         ' Now call the Windows API function again, but this time with
         ' the correctly sized string that receives the UTF-16 representation
         '
         ' The pointer to the VBA string *must* be calculated by calling "StrPtr"
         ' otherwise VBA will convert to ANSI before calling the API function
         '
         MultiByteToWideChar CP_UTF8, _
                             0&, _
                             utf8ByteArray(LBound(utf8ByteArray)), _
                             utf8Size, _
                             StrPtr(result), _
                             resultSize
      Else
         If Err.LastDllError = ERROR_NO_UNICODE_TRANSLATION Then
            Err.Raise ERR_NUM_INVALID_UTF8_CODE, _
                      ERR_STR_CLASS_NAME, _
                      ERR_STR_INVALID_UTF8_CODE
         Else
            HandleAPIError "MultiByteToWideChar"
         End If
      End If
   Else
      result = ""
   End If

   FromUTF8ToVB = result
End Function
