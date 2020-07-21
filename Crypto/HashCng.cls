VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "HashCng"
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
'|
'|-------------------------------------------------------------------------
'| Class               | HashCng
'|---------------------+---------------------------------------------------
'| Description         | Universal hasher using CNG
'|---------------------+---------------------------------------------------
'| Author              | Frank Schwab
'|---------------------+---------------------------------------------------
'| Version             | 1.0.0
'|---------------------+---------------------------------------------------
'| Changes             | 2020-07-18  Created. fhs
'|---------------------+---------------------------------------------------
'| Typical usage       | Dim hc256 As New HashCng
'|                     | hc256.InitializeSHA256
'|                     | hc256.DigestBytes aByteArray
'|                     | hc256.DigestStringAsUTF8 "Some string"
'|                     | Dim hashValue() as Byte
'|                     | hashValue = hc256.GetHash()
'|                     | ' Digest... and GetHash may be called
'|                     | ' again as often, as necessary
'|                     |
'|                     | ' If you want to change the hash type call
'|                     | ' any of the Initialize... methods
'|---------------------+---------------------------------------------------
'

Option Explicit

'
' Private constants for error messages
'
Private Const STR_ERROR_SOURCE   As String = "HashCng"
Private Const ERR_BASE As Long = vbObjectError + 33851

' Crypto API error
Private Const STR_ERR_BCRYPT_API As String = "Unable to %1. %2 returned code 0x%3: %4"
Private Const ERR_BCRYPT_API     As Long = ERR_BASE

' Error message when this object has not been used properly
Private Const STR_ERR_INVALID_USAGE As String = "Hash has not been %1"
Private Const ERR_INVALID_USAGE     As Long = ERR_BASE + 1

' Error message when key is too short
Private Const STR_ERR_KEY_TOO_SHORT As String = "Key is too short (minimum length: %1)"
Private Const ERR_KEY_TOO_SHORT     As Long = ERR_BASE + 2

'
' NT Status Code
'
Private Const STATUS_SUCCESS As Long = 0&

'
' CNG constants
'

' CNG Algorithm Identifiers
Private Const BCRYPT_SHA1_ALGORITHM   As String = "SHA1" & vbNullChar
Private Const BCRYPT_SHA256_ALGORITHM As String = "SHA256" & vbNullChar
Private Const BCRYPT_SHA384_ALGORITHM As String = "SHA384" & vbNullChar
Private Const BCRYPT_SHA512_ALGORITHM As String = "SHA512" & vbNullChar

' Cryptography Primitive Property Identifiers
Private Const BCRYPT_OBJECT_LENGTH As String = "ObjectLength" & vbNullChar
Private Const BCRYPT_HASH_LENGTH As String = "HashDigestLength" & vbNullChar

' Algorithm flags
Private Const BCRYPT_NO_FLAG              As Long = 0&
Private Const BCRYPT_ALG_HANDLE_HMAC_FLAG As Long = &H8&
Private Const BCRYPT_HASH_REUSABLE_FLAG   As Long = &H20&

'
' CNG API declarations
'
Private Declare Function BCryptOpenAlgorithmProvider _
                Lib "bcrypt.dll" _
                (ByRef phAlgorithm As Long, _
                 ByVal pszAlgId As Long, _
                 ByVal pszImplementation As Long, _
                 ByVal dwFlags As Long) _
                As Long

Private Declare Function BCryptGetProperty _
                Lib "bcrypt.dll" _
                (ByVal hObject As Long, _
                 ByVal pszProperty As Long, _
                 ByRef pbOutput As Any, _
                 ByVal cbOutput As Long, _
                 ByRef pcbResult As Long, _
                 ByVal dwFlags As Long) _
                As Long
                
Private Declare Function BCryptCreateHash _
                Lib "bcrypt.dll" _
                (ByVal hAlgorithm As Long, _
                 ByRef phHash As Long, _
                 ByVal pbHashObject As Long, _
                 ByVal cbHashObject As Long, _
                 ByVal pbSecret As Long, _
                 ByVal cbSecret As Long, _
                 ByVal dwFlags As Long) _
                As Long

Private Declare Function BCryptHashData _
                Lib "bcrypt.dll" _
                (ByVal hHash As Long, _
                 ByRef pbInput As Any, _
                 ByVal cbInput As Long, _
                 ByVal dwFlags As Long) _
                As Long

Private Declare Function BCryptFinishHash _
                Lib "bcrypt.dll" _
                (ByVal hHash As Long, _
                 ByRef pbOutput As Any, _
                 ByVal cbOutput As Long, _
                 ByVal dwFlags As Long) _
                As Long
                
Private Declare Function BCryptDestroyHash _
                Lib "bcrypt.dll" _
                (ByVal hHash As Long) _
                As Long
                
Private Declare Function BCryptCloseAlgorithmProvider _
                Lib "bcrypt.dll" _
                (ByVal hAlgorithm As Long, _
                 ByVal dwFlags As Long) _
                As Long

'
' Private check constants
'
Private Const MINIMUM_KEY_LENGTH As Long = 20

'
' Instance variables
'
Private m_MM As New MessageManager
Private m_U8C As New UTF8Converter

Private m_AlgorithmProviderHandle As Long
Private m_HashHandle As Long
Private m_HashSize As Long
Private m_IsHMAC As Boolean

'
' Private methods
'

'
'+--------------------------------------------------------------------------
'| Method           | IsNTSuccess
'|------------------+-------------------------------------------------------
'| Description      | Test if an NT status code means success
'|------------------+-------------------------------------------------------
'| Parameter        | ntStatusCode: NT status code to check
'|------------------+-------------------------------------------------------
'| Return values    | ./.
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2020-07-18  Created. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | ./.
'+--------------------------------------------------------------------------
'
Private Function IsNTSuccess(ByVal ntStatusCode As Long) As Boolean
   IsNTSuccess = (ntStatusCode >= 0)
End Function

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
'|------------------+-------------------------------------------------------
'| Remarks          | ./.
'+--------------------------------------------------------------------------
'
Private Function GetArrayLength(ByRef aByteArray() As Byte) As Long
   If (Not Not aByteArray) <> 0 Then  ' This is a hack! "Not Not anArray" is 0 if the array is empty and not 0 if it is not
      GetArrayLength = UBound(aByteArray) - LBound(aByteArray) + 1
   Else
      GetArrayLength = 0
   End If
End Function

'
'+--------------------------------------------------------------------------
'| Method           | CheckKey
'|------------------+-------------------------------------------------------
'| Description      | Check if key is long enough
'|------------------+-------------------------------------------------------
'| Parameter        | key: Key byte array to check
'|------------------+-------------------------------------------------------
'| Return values    | Length of key
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2020-07-18  Created. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | ./.
'+--------------------------------------------------------------------------
'
Private Function CheckKey(ByRef key() As Byte) As Long
   Dim arrayLength As Long
   
   arrayLength = GetArrayLength(key)

   If arrayLength < MINIMUM_KEY_LENGTH Then _
      Err.Raise ERR_KEY_TOO_SHORT, _
                STR_ERROR_SOURCE, _
                m_MM.FormatMessageWithParameters(STR_ERR_KEY_TOO_SHORT, Format$(MINIMUM_KEY_LENGTH))

   CheckKey = arrayLength
End Function

' Destroy methods

'
'+--------------------------------------------------------------------------
'| Method           | DestroyHash
'|------------------+-------------------------------------------------------
'| Description      | Destroy the hash object
'|------------------+-------------------------------------------------------
'| Parameter        | ./.
'|------------------+-------------------------------------------------------
'| Return values    | ./.
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2020-07-18  Created. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | ./.
'+--------------------------------------------------------------------------
'
Private Sub DestroyHash()
   If m_HashHandle <> 0 Then
      Dim rc As Long

      rc = BCryptDestroyHash(m_HashHandle)

      If IsNTSuccess(rc) Then
         m_HashHandle = 0
      Else
         Err.Raise ERR_BCRYPT_API, _
                   STR_ERROR_SOURCE, _
                   m_MM.FormatMessageWithParameters(STR_ERR_BCRYPT_API, _
                                                    "destroy hash", _
                                                    "BCryptDestroyHash", _
                                                    Hex$(rc), _
                                                    m_MM.GetMessageForNTStatusCode(rc))
      End If
   End If
End Sub

'
'+--------------------------------------------------------------------------
'| Method           | DestroyProvider
'|------------------+-------------------------------------------------------
'| Description      | Destroy the provider object
'|------------------+-------------------------------------------------------
'| Parameter        | ./.
'|------------------+-------------------------------------------------------
'| Return values    | ./.
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2020-07-18  Created. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | ./.
'+--------------------------------------------------------------------------
'
Private Sub DestroyProvider()
   DestroyHash

   If m_AlgorithmProviderHandle <> 0 Then
      Dim rc As Long

      rc = BCryptCloseAlgorithmProvider(m_AlgorithmProviderHandle, 0&)

      If IsNTSuccess(rc) Then
         m_AlgorithmProviderHandle = 0
      Else
         Err.Raise ERR_BCRYPT_API, _
                   STR_ERROR_SOURCE, _
                   m_MM.FormatMessageWithParameters(STR_ERR_BCRYPT_API, _
                                                    "close algorithm provider", _
                                                    "BCryptCloseAlgorithmProvider", _
                                                    Hex$(rc), _
                                                    m_MM.GetMessageForNTStatusCode(rc))
      End If
   End If
End Sub

' Initialize methods

'
'+--------------------------------------------------------------------------
'| Method           | InitializeProvider
'|------------------+-------------------------------------------------------
'| Description      | Initialize the provider object
'|------------------+-------------------------------------------------------
'| Parameter        | hashAlgorithm: The BCRYPT algorithm string
'|                  | algorithmFlag: A flag to use with the OpenAlgorithm
'|------------------+-------------------------------------------------------
'| Return values    | ./.
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2020-07-18  Created. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | A former provider will be destroyed
'+--------------------------------------------------------------------------
'
Private Sub InitializeProvider(ByRef hashAlgorithm As String, ByVal algorithmFlag As Long)
   DestroyProvider

   Dim rc As Long

   rc = BCryptOpenAlgorithmProvider(m_AlgorithmProviderHandle, _
                                    StrPtr(hashAlgorithm), _
                                    0&, _
                                    BCRYPT_HASH_REUSABLE_FLAG Or algorithmFlag)

   If Not IsNTSuccess(rc) Then
      Err.Raise ERR_BCRYPT_API, _
             STR_ERROR_SOURCE, _
             m_MM.FormatMessageWithParameters(STR_ERR_BCRYPT_API, _
                                              "open algorithm provider", _
                                              "BCryptOpenAlgorithmProvider", _
                                              Hex$(rc), _
                                              m_MM.GetMessageForNTStatusCode(rc))
   End If
End Sub

'
'+--------------------------------------------------------------------------
'| Method           | SetHashSize
'|------------------+-------------------------------------------------------
'| Description      | Set the hash size form the hash object
'|------------------+-------------------------------------------------------
'| Parameter        | ./.
'|------------------+-------------------------------------------------------
'| Return values    | ./.
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2020-07-18  Created. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | ./.
'+--------------------------------------------------------------------------
'
Private Sub SetHashSize()
   Dim rc As Long

   Dim hashSizeSize As Long

   hashSizeSize = LenB(m_HashSize)

   Dim newHashSize As Long

   rc = BCryptGetProperty(m_HashHandle, StrPtr(BCRYPT_HASH_LENGTH), newHashSize, hashSizeSize, hashSizeSize, 0&)

   If IsNTSuccess(rc) Then
      If newHashSize <> m_HashSize Then _
         m_HashSize = newHashSize
   Else
      Err.Raise ERR_BCRYPT_API, _
                STR_ERROR_SOURCE, _
                m_MM.FormatMessageWithParameters(STR_ERR_BCRYPT_API, _
                                                 "get hash size", _
                                                 "BCryptGetProperty", _
                                                 Hex$(rc), _
                                                 m_MM.GetMessageForNTStatusCode(rc))
   End If
End Sub

'
'+--------------------------------------------------------------------------
'| Method           | InitializeHash
'|------------------+-------------------------------------------------------
'| Description      | Initialize the hash object for a hash
'|------------------+-------------------------------------------------------
'| Parameter        | hashAlgorithm: The BCRYPT algorithm string
'|------------------+-------------------------------------------------------
'| Return values    | ./.
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2020-07-18  Created. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | Initializes the provider, as well
'+--------------------------------------------------------------------------
'
Private Sub InitializeHash(ByRef hashAlgorithm As String)
   InitializeProvider hashAlgorithm, BCRYPT_NO_FLAG
   
   Dim rc As Long

   rc = BCryptCreateHash(m_AlgorithmProviderHandle, m_HashHandle, 0&, 0&, 0&, 0&, BCRYPT_HASH_REUSABLE_FLAG)

   If IsNTSuccess(rc) Then
      SetHashSize
   Else
      Err.Raise ERR_BCRYPT_API, _
                STR_ERROR_SOURCE, _
                m_MM.FormatMessageWithParameters(STR_ERR_BCRYPT_API, _
                                                 "create hash", _
                                                 "BCryptCreateHash", _
                                                 Hex$(rc), _
                                                 m_MM.GetMessageForNTStatusCode(rc))
   End If

   m_IsHMAC = False
End Sub

'
'+--------------------------------------------------------------------------
'| Method           | InitializeHMAC
'|------------------+-------------------------------------------------------
'| Description      | Initialize the hash object for an HMAC
'|------------------+-------------------------------------------------------
'| Parameter        | hashAlgorithm: The BCRYPT algorithm string
'|------------------+-------------------------------------------------------
'| Return values    | ./.
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2020-07-18  Created. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | Initializes the provider, as well
'+--------------------------------------------------------------------------
'
Private Sub InitializeHMAC(ByRef hashAlgorithm As String, ByRef key() As Byte)
   InitializeProvider hashAlgorithm, BCRYPT_ALG_HANDLE_HMAC_FLAG

   Dim keyLength As Long

   keyLength = CheckKey(key)

   Dim keyPointer As Long
   
   keyPointer = VarPtr(key(LBound(key)))
   
   Dim rc As Long

   rc = BCryptCreateHash(m_AlgorithmProviderHandle, m_HashHandle, 0&, 0&, keyPointer, keyLength, BCRYPT_HASH_REUSABLE_FLAG)

   If IsNTSuccess(rc) Then
      SetHashSize
   Else
      Err.Raise ERR_BCRYPT_API, _
                STR_ERROR_SOURCE, _
                m_MM.FormatMessageWithParameters(STR_ERR_BCRYPT_API, _
                                                 "create hash", _
                                                 "BCryptCreateHash", _
                                                 Hex$(rc), _
                                                 m_MM.GetMessageForNTStatusCode(rc))
   End If

   m_IsHMAC = True
End Sub

'
'+--------------------------------------------------------------------------
'| Method           | CheckState
'|------------------+-------------------------------------------------------
'| Description      | Check if hasher is in a valid state
'|------------------+-------------------------------------------------------
'| Parameter        | ./.
'|------------------+-------------------------------------------------------
'| Return values    | ./.
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2020-07-18  Created. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | Raises error if hasher has not been initialized
'+--------------------------------------------------------------------------
'
Private Sub CheckState()
   If m_HashHandle = 0 Then _
      Err.Raise ERR_INVALID_USAGE, _
                STR_ERROR_SOURCE, _
                m_MM.FormatMessageWithParameters(STR_ERR_INVALID_USAGE, "initialized")
End Sub

'
' Public properties
'

'
'+--------------------------------------------------------------------------
'| Method           | HashSize
'|------------------+-------------------------------------------------------
'| Description      | Get size of hash in bytes
'|------------------+-------------------------------------------------------
'| Parameter        | ./.
'|------------------+-------------------------------------------------------
'| Return values    | Size of hash in bytes
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2020-07-18  Created. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | Raises error if hasher has not been initialized
'+--------------------------------------------------------------------------
'
Public Property Get HashSize() As Long
   CheckState

   HashSize = m_HashSize
End Property

'
'+--------------------------------------------------------------------------
'| Method           | IsHMAC
'|------------------+-------------------------------------------------------
'| Description      | Is current hash a HMAC?
'|------------------+-------------------------------------------------------
'| Parameter        | ./.
'|------------------+-------------------------------------------------------
'| Return values    | True, if current hash is an HMAC, false, if not
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2020-07-18  Created. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | Raises error if hasher has not been initialized
'+--------------------------------------------------------------------------
'
Public Property Get IsHMAC() As Boolean
   CheckState

   IsHMAC = m_IsHMAC
End Property

'
' Public methods
'

' Initialization methods

'
'+--------------------------------------------------------------------------
'| Method           | InitializeSHA1Hash
'|------------------+-------------------------------------------------------
'| Description      | Initialize hasher to use an SHA-1 hash
'|------------------+-------------------------------------------------------
'| Parameter        | ./.
'|------------------+-------------------------------------------------------
'| Return values    | ./.
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2020-07-18  Created. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | ./.
'+--------------------------------------------------------------------------
'
Public Sub InitializeSHA1Hash()
   InitializeHash BCRYPT_SHA1_ALGORITHM
End Sub

'
'+--------------------------------------------------------------------------
'| Method           | InitializeSHA1HMAC
'|------------------+-------------------------------------------------------
'| Description      | Initialize hasher to use an SHA-1 HMAC
'|------------------+-------------------------------------------------------
'| Parameter        | key: Key to use for HMAC
'|------------------+-------------------------------------------------------
'| Return values    | ./.
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2020-07-18  Created. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | ./.
'+--------------------------------------------------------------------------
'
Public Sub InitializeSHA1HMAC(ByRef key() As Byte)
   InitializeHMAC BCRYPT_SHA1_ALGORITHM, key
End Sub

'
'+--------------------------------------------------------------------------
'| Method           | InitializeSHA256Hash
'|------------------+-------------------------------------------------------
'| Description      | Initialize hasher to use an SHA-2-256 hash
'|------------------+-------------------------------------------------------
'| Parameter        | ./.
'|------------------+-------------------------------------------------------
'| Return values    | ./.
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2020-07-18  Created. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | ./.
'+--------------------------------------------------------------------------
'
Public Sub InitializeSHA256Hash()
   InitializeHash BCRYPT_SHA256_ALGORITHM
End Sub

'
'+--------------------------------------------------------------------------
'| Method           | InitializeSHA256HMAC
'|------------------+-------------------------------------------------------
'| Description      | Initialize hasher to use an SHA-2-256 HMAC
'|------------------+-------------------------------------------------------
'| Parameter        | key: Key to use for HMAC
'|------------------+-------------------------------------------------------
'| Return values    | ./.
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2020-07-18  Created. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | ./.
'+--------------------------------------------------------------------------
'
Public Sub InitializeSHA256HMAC(ByRef key() As Byte)
   InitializeHMAC BCRYPT_SHA256_ALGORITHM, key
End Sub

'
'+--------------------------------------------------------------------------
'| Method           | InitializeSHA384Hash
'|------------------+-------------------------------------------------------
'| Description      | Initialize hasher to use an SHA-2-384 hash
'|------------------+-------------------------------------------------------
'| Parameter        | ./.
'|------------------+-------------------------------------------------------
'| Return values    | ./.
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2020-07-18  Created. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | ./.
'+--------------------------------------------------------------------------
'
Public Sub InitializeSHA384Hash()
   InitializeHash BCRYPT_SHA384_ALGORITHM
End Sub

'
'+--------------------------------------------------------------------------
'| Method           | InitializeSHA384HMAC
'|------------------+-------------------------------------------------------
'| Description      | Initialize hasher to use an SHA-2-384 HMAC
'|------------------+-------------------------------------------------------
'| Parameter        | key: Key to use for HMAC
'|------------------+-------------------------------------------------------
'| Return values    | ./.
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2020-07-18  Created. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | ./.
'+--------------------------------------------------------------------------
'
Public Sub InitializeSHA384HMAC(ByRef key() As Byte)
   InitializeHMAC BCRYPT_SHA384_ALGORITHM, key
End Sub

'
'+--------------------------------------------------------------------------
'| Method           | InitializeSHA512Hash
'|------------------+-------------------------------------------------------
'| Description      | Initialize hasher to use an SHA-2-512 hash
'|------------------+-------------------------------------------------------
'| Parameter        | ./.
'|------------------+-------------------------------------------------------
'| Return values    | ./.
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2020-07-18  Created. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | ./.
'+--------------------------------------------------------------------------
'
Public Sub InitializeSHA512Hash()
   InitializeHash BCRYPT_SHA512_ALGORITHM
End Sub

'
'+--------------------------------------------------------------------------
'| Method           | InitializeSHA512HMAC
'|------------------+-------------------------------------------------------
'| Description      | Initialize hasher to use an SHA-2-512 HMAC
'|------------------+-------------------------------------------------------
'| Parameter        | key: Key to use for HMAC
'|------------------+-------------------------------------------------------
'| Return values    | ./.
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2020-07-18  Created. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | ./.
'+--------------------------------------------------------------------------
'
Public Sub InitializeSHA512HMAC(ByRef key() As Byte)
   InitializeHMAC BCRYPT_SHA512_ALGORITHM, key
End Sub

' Digest methods

'
'+--------------------------------------------------------------------------
'| Method           | DigestBytes
'|------------------+-------------------------------------------------------
'| Description      | Hashes a byte array
'|------------------+-------------------------------------------------------
'| Parameter        | sourceByteArray: Source bytes to hash
'|------------------+-------------------------------------------------------
'| Return values    | ./.
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2020-07-18  Created. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | ./.
'+--------------------------------------------------------------------------
'
Public Sub DigestBytes(ByRef sourceByteArray() As Byte)
   CheckState
   
   Dim arrayLength As Long

   arrayLength = GetArrayLength(sourceByteArray)

   If arrayLength > 0 Then
      Dim rc As Long

      rc = BCryptHashData(m_HashHandle, sourceByteArray(LBound(sourceByteArray)), arrayLength, 0&)

      If Not IsNTSuccess(rc) Then _
         Err.Raise ERR_BCRYPT_API, _
                   STR_ERROR_SOURCE, _
                   m_MM.FormatMessageWithParameters(STR_ERR_BCRYPT_API, _
                                                    "get hash", _
                                                    "BCryptHashData", _
                                                    Hex$(rc), _
                                                    m_MM.GetMessageForNTStatusCode(rc))
   End If
End Sub

'
'+--------------------------------------------------------------------------
'| Method           | DigestStringAsAnsi
'|------------------+-------------------------------------------------------
'| Description      | Hashes a string and uses Windows ANSI character code
'|                  | conversion from characters to bytes
'|------------------+-------------------------------------------------------
'| Parameter        | sourceString: Source string to hash
'|------------------+-------------------------------------------------------
'| Return values    | ./.
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2020-07-18  Created. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | ./.
'+--------------------------------------------------------------------------
'
Public Sub DigestStringAsANSI(ByRef sourceString As String)
   If Len(sourceString) > 0 Then
      Dim ansiBytesOfString() As Byte

      ansiBytesOfString = StrConv(sourceString, vbFromUnicode)
      Me.DigestBytes ansiBytesOfString
   End If
End Sub

'
'+--------------------------------------------------------------------------
'| Method           | DigestStringAsUTF8
'|------------------+-------------------------------------------------------
'| Description      | Hashes a string and uses UTF-8 character code
'|                  | conversion from characters to bytes
'|------------------+-------------------------------------------------------
'| Parameter        | sourceString: Source string to hash
'|------------------+-------------------------------------------------------
'| Return values    | ./.
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2020-07-18  Created. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | ./.
'+--------------------------------------------------------------------------
'
Public Sub DigestStringAsUTF8(ByRef sourceString As String)
   If Len(sourceString) > 0 Then
      Dim utf8BytesOfString() As Byte

      utf8BytesOfString = m_U8C.FromVBToUTF8(sourceString)
      Me.DigestBytes utf8BytesOfString
   End If
End Sub

' Get method

'
'+--------------------------------------------------------------------------
'| Method           | GetHash
'|------------------+-------------------------------------------------------
'| Description      | Get hash value
'|------------------+-------------------------------------------------------
'| Parameter        | ./.
'|------------------+-------------------------------------------------------
'| Return values    | Hash value
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2020-07-18  Created. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | After a call to this method the hash is reset to
'|                  | start again from the beginning
'+--------------------------------------------------------------------------
'
Public Function GetHash() As Byte()
   CheckState

   Dim rc As Long

   Dim result() As Byte
   ReDim result(1 To m_HashSize)
   
   rc = BCryptFinishHash(m_HashHandle, result(1), m_HashSize, 0&)
   
   If IsNTSuccess(rc) Then
      GetHash = result
   Else
      Err.Raise ERR_BCRYPT_API, _
                STR_ERROR_SOURCE, _
                m_MM.FormatMessageWithParameters(STR_ERR_BCRYPT_API, _
                                                 "finish hash", _
                                                 "BCryptFinishHash", _
                                                 Hex$(rc), _
                                                 m_MM.GetMessageForNTStatusCode(rc))
   End If
End Function

'
' Class methods
'
Private Sub Class_Terminate()
   DestroyProvider   ' Destroy everything on termination to prevent resource leaks
End Sub
