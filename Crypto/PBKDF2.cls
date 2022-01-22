VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PBKDF2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
'+-------------------------------------------------------------------------
'|
'| SPDX-FileCopyrightText: 2019 Frank Schwab
'|
'| SPDX-License-Identifier: MIT
'|
'| Copyright 2019, Frank Schwab
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
'| Class               | PBKDF2
'|---------------------+---------------------------------------------------
'| Description         | Implements the PBKDF2 key derivation function
'|                     | as defined in RFC 2898 using the Microsoft
'|                     | Crypto API
'|---------------------+---------------------------------------------------
'| Author              | Frank Schwab
'|---------------------+---------------------------------------------------
'| Version             | 1.1.0
'|---------------------+---------------------------------------------------
'| Changes             | 2019-02-19  Created. fhs
'|                     | 2022-01-22  Made 64 bit compatible. fhs
'|---------------------+---------------------------------------------------
'| Typical usage       | Dim sha1 As New SHA1Hash
'|                     | sha1.Initialize
'|                     | sha1.DigestString(strPassword)
'|                     | sha1.Finish
'|                     | hashPW = sha1.HashValueAsBase64String
'|                     |
'|                     | or, if you are lazy ;-)
'|                     |
'|                     | sha1.DigestString(strPassword)
'|                     | hashPW = sha1.HashValueAsBase64String
'+-------------------------------------------------------------------------
'

Option Explicit


'
' Private constantds
'

' Algorithm flags
Private Const BCRYPT_ALG_HANDLE_HMAC_FLAG As Long = 8

' Algorithm ids
Private Const BCRYPT_MD4_ALGORITHM As String = "MD4"
Private Const BCRYPT_MD5_ALGORITHM As String = "MD5"
Private Const BCRYPT_SHA1_ALGORITHM As String = "SHA1"
Private Const BCRYPT_SHA256_ALGORITHM As String = "SHA256"
Private Const BCRYPT_SHA384_ALGORITHM As String = "SHA384"
Private Const BCRYPT_SHA512_ALGORITHM As String = "SHA512"

Private Const BCRYPT_PBKDF2_ALGORITHM As String = "PBKDF2"

Private Const BCRYPT_HASH_LENGTH As String = "HashDigestLength"

' Failed/Succeeded
Private Const STATUS_SUCCESS As Long = 0


'
' Constants for error messages
'
Private Const STR_ERROR_SOURCE   As String = "PBKDF2"
Private Const ERR_BASE As Long = vbObjectError + 21234

' Crypto API error
Private Const STR_ERR_CRYPTO_API As String = "CNG Crypto API function %1 returned code 0x%2: %3"
Private Const ERR_CRYPTO_API     As Long = ERR_BASE


'
' API declarations
'
Private Declare PtrSafe Function BCryptOpenAlgorithmProvider Lib "bcrypt" ( _
   ByRef phAlgorithm As LongPtr, _
   ByVal pszAlgId As LongPtr, _
   ByVal pszImplementation As LongPtr, _
   ByVal dwFlags As Long) As Long

Private Declare PtrSafe Function BCryptCloseAlgorithmProvider Lib "bcrypt" ( _
   ByVal hAlgorithm As LongPtr, _
   ByVal dwFlags As Long) As Long

Private Declare PtrSafe Function BCryptGetProperty Lib "bcrypt" ( _
   ByVal hObject As LongPtr, _
   ByVal pszProperty As LongPtr, _
   ByVal pbOutput As LongPtr, _
   ByVal cbOutput As Long, _
   ByRef cbResult As Long, _
   ByVal dwFlags As Long) As Long

Private Declare PtrSafe Function BCryptDeriveKeyPBKDF2 Lib "bcrypt" ( _
   ByVal hPrf As LongPtr, _
   ByRef pbPassword As Byte, _
   ByVal cbPassword As Long, _
   ByRef pbSalt As Byte, _
   ByVal cbSalt As Long, _
   ByVal cIterations As LongLong, _
   ByRef pbDerivedKey As Byte, _
   ByVal cbDerivedKey As Long, _
   ByVal dwFlags As Long) As Long

' dwDummy is a dummy parameter because cIterations is a LongLong
' and the LongLong is simulated by a long followed by a dummy long
Private Declare PtrSafe Function XBCryptDeriveKeyPBKDF2 Lib "bcrypt" ( _
   ByVal hPrf As LongPtr, _
   ByRef pbPassword As Byte, _
   ByVal cbPassword As Long, _
   ByRef pbSalt As Byte, _
   ByVal cbSalt As Long, _
   ByVal cIterations As Long, _
   ByVal dwDummy As Long, _
   ByRef pbDerivedKey As Byte, _
   ByVal cbDerivedKey As Long, _
   ByVal dwFlags As Long) As Long


'
' Private attributes
'

'
' Private Data
'

'
' Instance variables
'
Private m_U8C As New UTF8Converter

'
' Private methods
'

'
'+--------------------------------------------------------------------------
'| Method           | GetU8C
'|------------------+-------------------------------------------------------
'| Description      | Get an UTF8 converter
'|------------------+-------------------------------------------------------
'| Parameter        | ./.
'|------------------+-------------------------------------------------------
'| Return values    | An UTF8 converter
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2019-02-19  Created. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | ./.
'+--------------------------------------------------------------------------
'
Private Function GetU8C() As UTF8Converter
   If m_U8C Is Nothing Then _
      Set m_U8C = New UTF8Converter

   Set GetU8C = m_U8C
End Function

'
'+--------------------------------------------------------------------------
'| Method           | GetLengthOfByteArray
'|------------------+-------------------------------------------------------
'| Description      | Get length of byte array
'|------------------+-------------------------------------------------------
'| Parameter        | aByteArray: Byte array
'|------------------+-------------------------------------------------------
'| Return values    | Length of byte array
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2019-02-19  Created. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | ./.
'+--------------------------------------------------------------------------
'
Private Function GetLengthOfByteArray(ByRef aByteArray() As Byte) As Long
   GetLengthOfByteArray = UBound(aByteArray) - LBound(aByteArray) + 1
End Function

'
'+--------------------------------------------------------------------------
'| Method           | ClearByteArray
'|------------------+-------------------------------------------------------
'| Description      | Clear a byte array
'|------------------+-------------------------------------------------------
'| Parameter        | aByteArray: Byte array to clear
'|------------------+-------------------------------------------------------
'| Return values    | ./.
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2019-02-19  Created. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | ./.
'+--------------------------------------------------------------------------
'
Private Sub ClearByteArray(ByRef aByteArray() As Byte)
   Dim i As Long
   
   For i = LBound(aByteArray) To UBound(aByteArray)
      aByteArray(i) = 0
   Next i
End Sub

'
'+--------------------------------------------------------------------------
'| Method           | HandleError
'|------------------+-------------------------------------------------------
'| Description      | Handle BCRYPT errors by throwing an exception
'|------------------+-------------------------------------------------------
'| Parameter        | apiFunctionName: Name of the failing API function
'|                  | rc: Status code of the failing API function
'|------------------+-------------------------------------------------------
'| Return values    | ./.
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2019-02-19  Created. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | This methods always throws an exception
'+--------------------------------------------------------------------------
'
Private Sub HandleError(ByRef apiFunctionName As String, _
                        ByVal rc As Long)
   Dim mm As New MessageManager
   
   Err.Raise ERR_CRYPTO_API, _
             STR_ERROR_SOURCE, _
             mm.FormatMessageWithParameters(STR_ERR_CRYPTO_API, _
                                            apiFunctionName, _
                                            Hex$(rc), _
                                            mm.GetMessageForNTStatusCode(rc))
End Sub

'
'+--------------------------------------------------------------------------
'| Method           | Pbkdf2WithHash
'|------------------+-------------------------------------------------------
'| Description      | Calculates a PBKDF2 hash.
'|------------------+-------------------------------------------------------
'| Parameter        | hashHandle:     Handle to a hash function
'|                  | passwordBytes:  The bytes to hash
'|                  | saltBytes:      The bytes of the salt
'|                  | iterationCount: The iteration count to use
'|------------------+-------------------------------------------------------
'| Return values    | The PBKDF2 hash as a byte array.
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2019-02-19  Created. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | ./.
'+--------------------------------------------------------------------------
'
Private Function Pbkdf2WithHash(ByVal hashHandle As LongPtr, _
                                ByRef passwordBytes() As Byte, _
                                ByRef saltBytes() As Byte, _
                                ByVal iterationCount As Long) As Byte()
   Dim dummy As Long

   Dim hashLength As Long

   Dim rc As Long

   rc = BCryptGetProperty(hashHandle, _
                          StrPtr(BCRYPT_HASH_LENGTH), _
                          VarPtr(hashLength), _
                          4, _
                          dummy, _
                          0)

   If rc <> STATUS_SUCCESS Then _
      HandleError "BCryptGetProperty", rc

   Dim result() As Byte

   ReDim result(1 To hashLength)

   Dim passwordLength As Long
   passwordLength = GetLengthOfByteArray(passwordBytes)

   Dim saltLength As Long
   saltLength = GetLengthOfByteArray(saltBytes)

   rc = BCryptDeriveKeyPBKDF2(hashHandle, _
                              passwordBytes(LBound(passwordBytes)), _
                              passwordLength, _
                              saltBytes(LBound(saltBytes)), _
                              saltLength, _
                              iterationCount, _
                              result(LBound(result)), _
                              hashLength, _
                              0)

   If rc <> STATUS_SUCCESS Then _
      HandleError "BCryptDeriveKeyPBKDF2", rc

   Pbkdf2WithHash = result
End Function

'
'+--------------------------------------------------------------------------
'| Method           | GetHandleForHash
'|------------------+-------------------------------------------------------
'| Description      | Get a handle for a hash function.
'|------------------+-------------------------------------------------------
'| Parameter        | hashAlgorithm: The hash algorithm name
'|------------------+-------------------------------------------------------
'| Return values    | The handle to the hash algorithm.
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2019-02-19  Created. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | ./.
'+--------------------------------------------------------------------------
'
Private Function GetHandleForHash(ByRef hashAlgorithm As String) As LongPtr
   Dim result As LongPtr

   Dim rc As Long

   rc = BCryptOpenAlgorithmProvider(result, _
                                    StrPtr(hashAlgorithm), _
                                    0&, _
                                    BCRYPT_ALG_HANDLE_HMAC_FLAG)
   If rc <> STATUS_SUCCESS Then _
      HandleError "BCryptOpenAlgorithmProvider", rc

   GetHandleForHash = result
End Function

'
'+--------------------------------------------------------------------------
'| Method           | Pbkdf2Wrapper
'|------------------+-------------------------------------------------------
'| Description      | Calculates a PBKDF2 hash.
'|------------------+-------------------------------------------------------
'| Parameter        | hashAlgorithm:  Algorithm name of a hash function
'|                  | passwordBytes:  The bytes to hash
'|                  | saltBytes:      The bytes of the salt
'|                  | iterationCount: The iteration count to use
'|------------------+-------------------------------------------------------
'| Return values    | The PBKDF2 hash as a byte array.
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2019-02-19  Created. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | This method first converts the password string to
'|                  | a byte array using the UTF8 encoding. Then it calculates
'|                  | the PBKDF2 hash values using these bytes.
'+--------------------------------------------------------------------------
'
Private Function Pbkdf2Wrapper(ByVal hashAlgorithm As String, _
                               ByRef password As String, _
                               ByRef saltBytes() As Byte, _
                               ByVal iterationCount As Long) As Byte()
   Dim passwordBytes() As Byte

   passwordBytes = GetU8C.FromVBToUTF8(password)

   Dim hashHandle As LongPtr
   
   hashHandle = GetHandleForHash(hashAlgorithm)
   
   Pbkdf2Wrapper = Pbkdf2WithHash(hashHandle, _
                                  passwordBytes, _
                                  saltBytes, _
                                  iterationCount)
   ClearByteArray passwordBytes

   Dim rc As Long

   rc = BCryptCloseAlgorithmProvider(hashHandle, _
                                     0&)

   If rc <> STATUS_SUCCESS Then _
      HandleError "BCryptCloseAlgorithmProvider", rc
End Function

'
' Public methods
'

'
'+--------------------------------------------------------------------------
'| Method           | Pbkdf2WithMD4
'|------------------+-------------------------------------------------------
'| Description      | Calculates a PBKDF2 hash with MD4 as the hash function.
'|------------------+-------------------------------------------------------
'| Parameter        | password:       Password to hsah
'|                  | saltBytes:      The bytes of the salt
'|                  | iterationCount: The iteration count to use
'|------------------+-------------------------------------------------------
'| Return values    | The PBKDF2 hash as a byte array.
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2019-02-19  Created. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | ./.
'+--------------------------------------------------------------------------
'
Public Function Pbkdf2WithMD4(ByRef password As String, _
                              ByRef saltBytes() As Byte, _
                              ByVal iterationCount As Long) As Byte()
   Pbkdf2WithMD4 = Pbkdf2Wrapper(BCRYPT_MD4_ALGORITHM, _
                                 password, _
                                 saltBytes, _
                                 iterationCount)
End Function

'
'+--------------------------------------------------------------------------
'| Method           | Pbkdf2WithMD5
'|------------------+-------------------------------------------------------
'| Description      | Calculates a PBKDF2 hash with MD5 as the hash function.
'|------------------+-------------------------------------------------------
'| Parameter        | password:       Password to hsah
'|                  | saltBytes:      The bytes of the salt
'|                  | iterationCount: The iteration count to use
'|------------------+-------------------------------------------------------
'| Return values    | The PBKDF2 hash as a byte array.
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2019-02-19  Created. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | ./.
'+--------------------------------------------------------------------------
'
Public Function Pbkdf2WithMD5(ByRef password As String, _
                              ByRef saltBytes() As Byte, _
                              ByVal iterationCount As Long) As Byte()
   Pbkdf2WithMD5 = Pbkdf2Wrapper(BCRYPT_MD5_ALGORITHM, _
                                 password, _
                                 saltBytes, _
                                 iterationCount)
End Function

'
'+--------------------------------------------------------------------------
'| Method           | Pbkdf2WithSHA1
'|------------------+-------------------------------------------------------
'| Description      | Calculates a PBKDF2 hash with SHA-1 as the hash function.
'|------------------+-------------------------------------------------------
'| Parameter        | password:       Password to hsah
'|                  | saltBytes:      The bytes of the salt
'|                  | iterationCount: The iteration count to use
'|------------------+-------------------------------------------------------
'| Return values    | The PBKDF2 hash as a byte array.
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2019-02-19  Created. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | ./.
'+--------------------------------------------------------------------------
'
Public Function Pbkdf2WithSHA1(ByRef password As String, _
                               ByRef saltBytes() As Byte, _
                               ByVal iterationCount As Long) As Byte()
   Pbkdf2WithSHA1 = Pbkdf2Wrapper(BCRYPT_SHA1_ALGORITHM, _
                                  password, _
                                  saltBytes, _
                                  iterationCount)
End Function

'
'+--------------------------------------------------------------------------
'| Method           | Pbkdf2WithSHA256
'|------------------+-------------------------------------------------------
'| Description      | Calculates a PBKDF2 hash with SHA-2-256 as the hash function.
'|------------------+-------------------------------------------------------
'| Parameter        | password:       Password to hsah
'|                  | saltBytes:      The bytes of the salt
'|                  | iterationCount: The iteration count to use
'|------------------+-------------------------------------------------------
'| Return values    | The PBKDF2 hash as a byte array.
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2019-02-19  Created. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | ./.
'+--------------------------------------------------------------------------
'
Public Function Pbkdf2WithSHA256(ByRef password As String, _
                                 ByRef saltBytes() As Byte, _
                                 ByVal iterationCount As Long) As Byte()
   Pbkdf2WithSHA256 = Pbkdf2Wrapper(BCRYPT_SHA256_ALGORITHM, _
                                    password, _
                                    saltBytes, _
                                    iterationCount)
End Function

'
'+--------------------------------------------------------------------------
'| Method           | Pbkdf2WithSHA384
'|------------------+-------------------------------------------------------
'| Description      | Calculates a PBKDF2 hash with SHA-2-384 as the hash function.
'|------------------+-------------------------------------------------------
'| Parameter        | password:       Password to hsah
'|                  | saltBytes:      The bytes of the salt
'|                  | iterationCount: The iteration count to use
'|------------------+-------------------------------------------------------
'| Return values    | The PBKDF2 hash as a byte array.
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2019-02-19  Created. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | ./.
'+--------------------------------------------------------------------------
'
Public Function Pbkdf2WithSHA384(ByRef password As String, _
                                 ByRef saltBytes() As Byte, _
                                 ByVal iterationCount As Long) As Byte()
   Pbkdf2WithSHA384 = Pbkdf2Wrapper(BCRYPT_SHA384_ALGORITHM, _
                                    password, _
                                    saltBytes, _
                                    iterationCount)
End Function

'
'+--------------------------------------------------------------------------
'| Method           | Pbkdf2WithSHA512
'|------------------+-------------------------------------------------------
'| Description      | Calculates a PBKDF2 hash with SHA-2-512 as the hash function.
'|------------------+-------------------------------------------------------
'| Parameter        | password:       Password to hsah
'|                  | saltBytes:      The bytes of the salt
'|                  | iterationCount: The iteration count to use
'|------------------+-------------------------------------------------------
'| Return values    | The PBKDF2 hash as a byte array.
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2019-02-19  Created. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | ./.
'+--------------------------------------------------------------------------
'
Public Function Pbkdf2WithSHA512(ByRef password As String, _
                                 ByRef saltBytes() As Byte, _
                                 ByVal iterationCount As Long) As Byte()
   Pbkdf2WithSHA512 = Pbkdf2Wrapper(BCRYPT_SHA512_ALGORITHM, _
                                    password, _
                                    saltBytes, _
                                    iterationCount)
End Function

