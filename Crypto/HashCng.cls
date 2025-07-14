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
'| SPDX-FileCopyrightText: 2020-2025 Frank Schwab
'|
'| SPDX-License-Identifier: MIT
'|
'| Copyright 2020-2025, Frank Schwab
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
'| Class               | HashCng
'|---------------------+---------------------------------------------------
'| Description         | Universal hasher using CNG
'|---------------------+---------------------------------------------------
'| Author              | Frank Schwab
'|---------------------+---------------------------------------------------
'| Version             | 1.3.0
'|---------------------+---------------------------------------------------
'| Changes             | 2020-07-18  Created. fhs
'|                     | 2020-08-03  Correct array length calculation
'|                     |             because it was a hack. fhs
'|                     | 2020-09-01  Made 64 bit compatible. fhs
'|                     | 2020-09-09  Simplified. fhs
'|                     | 2020-09-10  Removed unnecessary constant. fhs
'|                     | 2021-01-22  Removed unnecessary terminating
'|                     |             null characters. fhs
'|                     | 2022-02-04  Corrected handling of hash result. fhs
'|                     | 2022-02-04  Reverted last change, fixed bugs. fhs
'|                     | 2022-03-18  Added hash of a part of a byte array
'|                     |             and file content. fhs
'|                     | 2022-03-19  Use BinaryFileInputStream. fhs
'|                     | 2024-08-10  Better 64 bit compatibility. fhs
'|                     | 2025-07-14  Add SHA-3 hashes and HMACs, KMACs
'|                     |             and IsMAC property. fhs
'|---------------------+---------------------------------------------------
'| Typical usage       | ' For getting a hash value
'|                     | Dim hc As New HashCng
'|                     | hc.InitializeSHA256
'|                     | hc.DigestBytes aByteArray
'|                     | hc.DigestStringAsUTF8 someText
'|                     | Dim hashValue() as Byte
'|                     | hashValue = hc.GetHash()
'|                     | ' hc.Digest... and hc.GetHash may be called
'|                     | ' again as often, as necessary.
'|                     | ' One may even switch the hash type by calling
'|                     | ' one of the hc.Initialize... methods.
'|                     |
'|                     | ' For getting an HMAC value
'|                     | Dim hc As New HashCng
'|                     | hc.InitializeSHA256HMAC key
'|                     | hc.DigestBytes aByteArray
'|                     | hc.DigestStringAsUTF8 someText
'|                     | hashValue = hc.GetHash
'+---------------------+---------------------------------------------------
'

Option Explicit

'
' Private constants for error messages
'
Private Const ERR_STR_CLASS_NAME   As String = "HashCng"
Private Const ERR_NUM_BASE As Long = vbObjectError + 33851

' Crypto API error
Private Const ERR_STR_BCRYPT_API As String = "Unable to %1. '%2' returned code 0x%3: %4"
Private Const ERR_NUM_BCRYPT_API As Long = ERR_NUM_BASE

' Error message when this object has not been used properly
Private Const ERR_STR_INVALID_USAGE As String = "Hash has not been %1"
Private Const ERR_NUM_INVALID_USAGE As Long = ERR_NUM_BASE + 1

' Error message when key is too short
Private Const ERR_STR_KEY_TOO_SHORT As String = "Key is too short (minimum length: %1)"
Private Const ERR_NUM_KEY_TOO_SHORT As Long = ERR_NUM_BASE + 2

' Error message when key is too short
Private Const ERR_STR_OFFSET_COUNT_EXCEED_UPPER_BOUND As String = "Offset %1 and count %2 exceed upper array boundary %3"
Private Const ERR_NUM_OFFSET_COUNT_EXCEED_UPPER_BOUND As Long = ERR_NUM_BASE + 3

'
' CNG constants
'

'
' No need for terminating vbNullChars as BSTRs always have a terminating null character.
'

' CNG Algorithm Identifiers
Private Const BCRYPT_SHA1_ALGORITHM   As String = "SHA1"
Private Const BCRYPT_SHA256_ALGORITHM As String = "SHA256"
Private Const BCRYPT_SHA384_ALGORITHM As String = "SHA384"
Private Const BCRYPT_SHA512_ALGORITHM As String = "SHA512"
Private Const BCRYPT_SHA3_256_ALGORITHM As String = "SHA3-256"
Private Const BCRYPT_SHA3_384_ALGORITHM As String = "SHA3-384"
Private Const BCRYPT_SHA3_512_ALGORITHM As String = "SHA3-512"
Private Const BCRYPT_KMAC128_ALGORITHM As String = "KMAC128"
Private Const BCRYPT_KMAC256_ALGORITHM As String = "KMAC256"

' Cryptography Primitive Property Identifiers
Private Const BCRYPT_HASH_LENGTH As String = "HashDigestLength"
Private Const BCRYPT_CUSTOMIZATION_STRING As String = "CustomizationString"

' Algorithm flags
Private Const BCRYPT_NO_FLAG              As Long = 0&
Private Const BCRYPT_ALG_HANDLE_HMAC_FLAG As Long = &H8&
Private Const BCRYPT_HASH_REUSABLE_FLAG   As Long = &H20&

'
' CNG API declarations
'
Private Declare PtrSafe Function BCryptOpenAlgorithmProvider _
                Lib "bcrypt.dll" _
                (ByRef phAlgorithm As LongPtr, _
                 ByVal pszAlgId As LongPtr, _
                 ByVal pszImplementation As LongPtr, _
                 ByVal dwFlags As Long) _
                As Long

Private Declare PtrSafe Function BCryptGetProperty _
                Lib "bcrypt.dll" _
                (ByVal hObject As LongPtr, _
                 ByVal pszProperty As LongPtr, _
                 ByRef pbOutput As Any, _
                 ByVal cbOutput As Long, _
                 ByRef pcbResult As Long, _
                 ByVal dwFlags As Long) _
                As Long

Private Declare PtrSafe Function BCryptSetProperty _
                Lib "bcrypt.dll" _
                (ByVal hObject As LongPtr, _
                 ByVal pszProperty As LongPtr, _
                 ByVal pbInput As LongPtr, _
                 ByVal cbInput As Long, _
                 ByVal dwFlags As Long) _
                As Long

Private Declare PtrSafe Function BCryptCreateHash _
                Lib "bcrypt.dll" _
                (ByVal hAlgorithm As LongPtr, _
                 ByRef phHash As LongPtr, _
                 ByVal pbHashObject As LongPtr, _
                 ByVal cbHashObject As Long, _
                 ByVal pbSecret As LongPtr, _
                 ByVal cbSecret As Long, _
                 ByVal dwFlags As Long) _
                As Long

Private Declare PtrSafe Function BCryptHashData _
                Lib "bcrypt.dll" _
                (ByVal hHash As LongPtr, _
                 ByRef pbInput As Byte, _
                 ByVal cbInput As Long, _
                 ByVal dwFlags As Long) _
                As Long

Private Declare PtrSafe Function BCryptFinishHash _
                Lib "bcrypt.dll" _
                (ByVal hHash As LongPtr, _
                 ByRef pbOutput As Byte, _
                 ByVal cbOutput As Long, _
                 ByVal dwFlags As Long) _
                As Long
                
Private Declare PtrSafe Function BCryptDestroyHash _
                Lib "bcrypt.dll" _
                (ByVal hHash As LongPtr) _
                As Long
                
Private Declare PtrSafe Function BCryptCloseAlgorithmProvider _
                Lib "bcrypt.dll" _
                (ByVal hAlgorithm As LongPtr, _
                 ByVal dwFlags As Long) _
                As Long

'
' Private check constants
'
Private Const MINIMUM_KEY_LENGTH As Long = 16

'
' Private file handling constants
'
Private Const BUFFER_SIZE As Long = 32768

'
' Instance variables
'
Private m_MM As New MessageManager
Private m_U8C As New UTF8Converter

Private m_AlgorithmProviderHandle As LongPtr
Private m_HashHandle As LongPtr
Private m_HashSize As Long
Private m_IsMAC As Boolean
Private m_IsHMAC As Boolean

'
' Private methods
'

'
'+--------------------------------------------------------------------------
'| Method           | IsNTSuccess
'|------------------+-------------------------------------------------------
'| Description      | Test if an NT status code means success.
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
'| Description      | Get length of a byte array.
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
'| Remarks          | ./.
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
'| Method           | HandleNTAPIError
'|------------------+-------------------------------------------------------
'| Description      | Handle Windows NT API error.
'|------------------+-------------------------------------------------------
'| Parameter        | actionDescription: Description of the action
'|                  | apiFunctionName  : Name of the API function
'|                  | ntStatusCode     : NT status code
'|------------------+-------------------------------------------------------
'| Return values    | Length of byte array
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2020-09-09  Created. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | This method does not return because it raises an error.
'+--------------------------------------------------------------------------
'
Private Sub HandleNTAPIError(ByRef actionText As String, ByRef apiFunctionName As String, ByRef ntStatusCode As Long)
   Err.Raise ERR_NUM_BCRYPT_API, _
             ERR_STR_CLASS_NAME, _
             m_MM.FormatMessageWithParameters(ERR_STR_BCRYPT_API, _
                                              actionText, _
                                              apiFunctionName, _
                                              Hex$(ntStatusCode), _
                                              m_MM.GetMessageForNTStatusCode(ntStatusCode))
End Sub

'
'+--------------------------------------------------------------------------
'| Method           | CheckKey
'|------------------+-------------------------------------------------------
'| Description      | Check if key is long enough.
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
      Err.Raise ERR_NUM_KEY_TOO_SHORT, _
                ERR_STR_CLASS_NAME, _
                m_MM.FormatMessageWithParameters(ERR_STR_KEY_TOO_SHORT, Format$(MINIMUM_KEY_LENGTH))

   CheckKey = arrayLength
End Function

' Destroy methods

'
'+--------------------------------------------------------------------------
'| Method           | DestroyHash
'|------------------+-------------------------------------------------------
'| Description      | Destroy the hash object.
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
   If m_HashHandle <> 0& Then
      Dim rc As Long

      rc = BCryptDestroyHash(m_HashHandle)

      If IsNTSuccess(rc) Then
         m_HashHandle = 0&
      Else
         HandleNTAPIError "destroy hash", "BCryptDestroyHash", rc
      End If
   End If
End Sub

'
'+--------------------------------------------------------------------------
'| Method           | DestroyProvider
'|------------------+-------------------------------------------------------
'| Description      | Destroy the provider object.
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

   If m_AlgorithmProviderHandle <> 0& Then
      Dim rc As Long

      rc = BCryptCloseAlgorithmProvider(m_AlgorithmProviderHandle, 0&)

      If IsNTSuccess(rc) Then
         m_AlgorithmProviderHandle = 0&
      Else
         HandleNTAPIError "close algorithm provider", "BCryptCloseAlgorithmProvider", rc
      End If
   End If
End Sub

' Initialize methods

'
'+--------------------------------------------------------------------------
'| Method           | InitializeProvider
'|------------------+-------------------------------------------------------
'| Description      | Initialize the provider object.
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

   ' Note that we use the "reusable" flag here so the hash can be reused
   rc = BCryptOpenAlgorithmProvider(m_AlgorithmProviderHandle, _
                                    StrPtr(hashAlgorithm), _
                                    0&, _
                                    BCRYPT_HASH_REUSABLE_FLAG Or algorithmFlag)

   If Not IsNTSuccess(rc) Then _
      HandleNTAPIError "open algorithm provider", "BCryptOpenAlgorithmProvider", rc
End Sub

'
'+--------------------------------------------------------------------------
'| Method           | SetHashSize
'|------------------+-------------------------------------------------------
'| Description      | Set the hash size form the hash object.
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

   hashSizeSize = 4&   ' Need to use a magic constant here. A long has 4 bytes.

   Dim newHashSize As Long

   rc = BCryptGetProperty(m_HashHandle, StrPtr(BCRYPT_HASH_LENGTH), newHashSize, hashSizeSize, hashSizeSize, 0&)

   If IsNTSuccess(rc) Then
      If newHashSize <> m_HashSize Then _
         m_HashSize = newHashSize
   Else
      HandleNTAPIError "get hash size", "BCryptGetProperty", rc
   End If
End Sub

'
'+--------------------------------------------------------------------------
'| Method           | SetCustomizationString
'|------------------+-------------------------------------------------------
'| Description      | Set the customization string for KMACs.
'|------------------+-------------------------------------------------------
'| Parameter        | customizationString: Customization byte array
'|------------------+-------------------------------------------------------
'| Return values    | ./.
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2025-07-14  Created. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | ./.
'+--------------------------------------------------------------------------
'
Private Sub SetCustomizationString(ByRef customizationString() As Byte)
   Dim rc As Long

   Dim csl As Long
   csl = GetArrayLength(customizationString)
   
   If csl > 0 Then
      rc = BCryptSetProperty(m_HashHandle, StrPtr(BCRYPT_CUSTOMIZATION_STRING), VarPtr(customizationString(LBound(customizationString))), csl, 0&)

      If Not IsNTSuccess(rc) Then _
         HandleNTAPIError "set customization string", "BCryptSetProperty", rc
   End If
End Sub

'
'+--------------------------------------------------------------------------
'| Method           | InitializeHash
'|------------------+-------------------------------------------------------
'| Description      | Initialize the hash object for a hash.
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
      HandleNTAPIError "create hash", "BCryptCreateHash", rc
   End If

   m_IsHMAC = False
End Sub

'
'+--------------------------------------------------------------------------
'| Method           | InitializeMAC
'|------------------+-------------------------------------------------------
'| Description      | Initialize the hash object for a MAC.
'|------------------+-------------------------------------------------------
'| Parameter        | hashAlgorithm: The BCRYPT algorithm string
'|                  | key:           Key for HMAC
'|                  | forHMAC:       True, if an HMAC is initialized
'|------------------+-------------------------------------------------------
'| Return values    | ./.
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2025-07-14  Created. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | Initializes the provider, as well.
'+--------------------------------------------------------------------------
'
Private Sub InitializeMAC(ByRef hashAlgorithm As String, ByRef key() As Byte, ByVal forHMAC As Boolean)
   Dim bcryptFlag As Long
   If forHMAC Then
      bcryptFlag = BCRYPT_ALG_HANDLE_HMAC_FLAG
   Else
      bcryptFlag = BCRYPT_NO_FLAG
   End If
   
   InitializeProvider hashAlgorithm, bcryptFlag

   Dim keyLength As Long

   keyLength = CheckKey(key)

   Dim keyPointer As LongPtr
   
   keyPointer = VarPtr(key(LBound(key)))
   
   Dim rc As Long

   rc = BCryptCreateHash(m_AlgorithmProviderHandle, m_HashHandle, 0&, 0&, keyPointer, keyLength, BCRYPT_HASH_REUSABLE_FLAG)

   If IsNTSuccess(rc) Then
      SetHashSize
   Else
      HandleNTAPIError "create KMAC hash", "BCryptCreateHash", rc
   End If

   m_IsMAC = True
   m_IsHMAC = forHMAC
End Sub

'
'+--------------------------------------------------------------------------
'| Method           | InitializeHMAC
'|------------------+-------------------------------------------------------
'| Description      | Initialize the hash object for an HMAC.
'|------------------+-------------------------------------------------------
'| Parameter        | hashAlgorithm: The BCRYPT algorithm string
'|                  | key:           Key for HMAC
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
   InitializeMAC hashAlgorithm, key, True
End Sub

'
'+--------------------------------------------------------------------------
'| Method           | InitializeKMAC
'|------------------+-------------------------------------------------------
'| Description      | Initialize the hash object for a KMAC.
'|------------------+-------------------------------------------------------
'| Parameter        | hashAlgorithm:       The BCRYPT algorithm string
'|                  | key:                 Key for KMAC
'|                  | customizationString: Customization byte array
'|------------------+-------------------------------------------------------
'| Return values    | ./.
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2025-07-14  Created. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | Initializes the provider, as well.
'+--------------------------------------------------------------------------
'
Private Sub InitializeKMAC(ByRef hashAlgorithm As String, ByRef key() As Byte, ByRef customizationString() As Byte)
   InitializeMAC hashAlgorithm, key, False
   SetCustomizationString customizationString
End Sub

'
'+--------------------------------------------------------------------------
'| Method           | CheckState
'|------------------+-------------------------------------------------------
'| Description      | Check if hasher is in a valid state.
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
      Err.Raise ERR_NUM_INVALID_USAGE, _
                ERR_STR_CLASS_NAME, _
                m_MM.FormatMessageWithParameters(ERR_STR_INVALID_USAGE, "initialized")
End Sub

'
' Public properties
'

'
'+--------------------------------------------------------------------------
'| Method           | HashSize
'|------------------+-------------------------------------------------------
'| Description      | Get size of hash in bytes.
'|------------------+-------------------------------------------------------
'| Parameter        | ./.
'|------------------+-------------------------------------------------------
'| Return values    | Size of hash in bytes.
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
'| Return values    | True, if current hash is an HMAC, false, if not.
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2020-07-18  Created. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | Raises error if hasher has not been initialized
'+--------------------------------------------------------------------------
'
Public Property Get isHMAC() As Boolean
   CheckState

   isHMAC = m_IsHMAC
End Property

'
'+--------------------------------------------------------------------------
'| Method           | IsMAC
'|------------------+-------------------------------------------------------
'| Description      | Is current hash a MAC?
'|------------------+-------------------------------------------------------
'| Parameter        | ./.
'|------------------+-------------------------------------------------------
'| Return values    | True, if current hash is an MAC, false, if not.
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2025-07-14  Created. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | Raises error if hasher has not been initialized
'+--------------------------------------------------------------------------
'
Public Property Get IsMAC() As Boolean
   CheckState

   IsMAC = m_IsMAC
End Property

'
' Public methods
'

' Initialization methods

'
'+--------------------------------------------------------------------------
'| Method           | InitializeSHA1Hash
'|------------------+-------------------------------------------------------
'| Description      | Initialize hasher to use an SHA-1 hash.
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
'| Description      | Initialize hasher to use an SHA-1 HMAC.
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
'| Description      | Initialize hasher to use an SHA-2-256 hash.
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
'| Description      | Initialize hasher to use an SHA-2-256 HMAC.
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
'| Description      | Initialize hasher to use an SHA-2-384 hash.
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
'| Description      | Initialize hasher to use an SHA-2-384 HMAC.
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
'| Description      | Initialize hasher to use an SHA-2-512 hash.
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
'| Description      | Initialize hasher to use an SHA-2-512 HMAC.
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

'
'+--------------------------------------------------------------------------
'| Method           | InitializeSHA3_256Hash
'|------------------+-------------------------------------------------------
'| Description      | Initialize hasher to use an SHA-3-256 hash.
'|------------------+-------------------------------------------------------
'| Parameter        | ./.
'|------------------+-------------------------------------------------------
'| Return values    | ./.
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2025-07-14  Created. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | Now, that Windows 11 finally supports SHA-3 9 years
'|                  | after it has been standardized by the NIST, we are
'|                  | able to offer it!
'+--------------------------------------------------------------------------
'
Public Sub InitializeSHA3_256Hash()
   InitializeHash BCRYPT_SHA3_256_ALGORITHM
End Sub

'
'+--------------------------------------------------------------------------
'| Method           | InitializeSHA3_256HMAC
'|------------------+-------------------------------------------------------
'| Description      | Initialize hasher to use an SHA-3-256 HMAC.
'|------------------+-------------------------------------------------------
'| Parameter        | key: Key to use for HMAC
'|------------------+-------------------------------------------------------
'| Return values    | ./.
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2025-07-14  Created. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | Now, that Windows 11 finally supports SHA-3 9 years
'|                  | after it has been standardized by the NIST, we are
'|                  | able to offer it!
'+--------------------------------------------------------------------------
'
Public Sub InitializeSHA3_256HMAC(ByRef key() As Byte)
   InitializeHMAC BCRYPT_SHA3_256_ALGORITHM, key
End Sub

'
'+--------------------------------------------------------------------------
'| Method           | InitializeSHA3_384Hash
'|------------------+-------------------------------------------------------
'| Description      | Initialize hasher to use an SHA-3-384 hash.
'|------------------+-------------------------------------------------------
'| Parameter        | ./.
'|------------------+-------------------------------------------------------
'| Return values    | ./.
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2025-07-14  Created. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | Now, that Windows 11 finally supports SHA-3 9 years
'|                  | after it has been standardized by the NIST, we are
'|                  | able to offer it!
'+--------------------------------------------------------------------------
'
Public Sub InitializeSHA3_384Hash()
   InitializeHash BCRYPT_SHA3_384_ALGORITHM
End Sub

'
'+--------------------------------------------------------------------------
'| Method           | InitializeSHA3_384HMAC
'|------------------+-------------------------------------------------------
'| Description      | Initialize hasher to use an SHA-3-384 HMAC.
'|------------------+-------------------------------------------------------
'| Parameter        | key: Key to use for HMAC
'|------------------+-------------------------------------------------------
'| Return values    | ./.
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2025-07-14  Created. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | Now, that Windows 11 finally supports SHA-3 9 years
'|                  | after it has been standardized by the NIST, we are
'|                  | able to offer it!
'+--------------------------------------------------------------------------
'
Public Sub InitializeSHA3_384HMAC(ByRef key() As Byte)
   InitializeHMAC BCRYPT_SHA3_384_ALGORITHM, key
End Sub

'
'+--------------------------------------------------------------------------
'| Method           | InitializeSHA3_512Hash
'|------------------+-------------------------------------------------------
'| Description      | Initialize hasher to use an SHA-3-512 hash.
'|------------------+-------------------------------------------------------
'| Parameter        | ./.
'|------------------+-------------------------------------------------------
'| Return values    | ./.
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2025-07-14  Created. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | Now, that Windows 11 finally supports SHA-3 9 years
'|                  | after it has been standardized by the NIST, we are
'|                  | able to offer it!
'+--------------------------------------------------------------------------
'
Public Sub InitializeSHA3_512Hash()
   InitializeHash BCRYPT_SHA3_512_ALGORITHM
End Sub

'
'+--------------------------------------------------------------------------
'| Method           | InitializeSHA3_512HMAC
'|------------------+-------------------------------------------------------
'| Description      | Initialize hasher to use an SHA-3-512 HMAC.
'|------------------+-------------------------------------------------------
'| Parameter        | key: Key to use for HMAC
'|------------------+-------------------------------------------------------
'| Return values    | ./.
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2025-07-14  Created. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | Now, that Windows 11 finally supports SHA-3 9 years
'|                  | after it has been standardized by the NIST, we are
'|                  | able to offer it!
'+--------------------------------------------------------------------------
'
Public Sub InitializeSHA3_512HMAC(ByRef key() As Byte)
   InitializeHMAC BCRYPT_SHA3_512_ALGORITHM, key
End Sub

'
'+--------------------------------------------------------------------------
'| Method           | InitializeKMAC128
'|------------------+-------------------------------------------------------
'| Description      | Initialize hasher to use an KMAC128 hash.
'|------------------+-------------------------------------------------------
'| Parameter        | key: Key to use for KMAC
'|                  | customizationString: Customization byte array
'|------------------+-------------------------------------------------------
'| Return values    | ./.
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2025-07-14  Created. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | Now, that Windows 11 finally supports KMAC 9 years
'|                  | after it has been standardized by the NIST, we are
'|                  | able to offer it!
'+--------------------------------------------------------------------------
'
Public Sub InitializeKMAC128(ByRef key() As Byte, ByRef customizationString() As Byte)
   InitializeKMAC BCRYPT_KMAC128_ALGORITHM, key, customizationString
End Sub

'
'+--------------------------------------------------------------------------
'| Method           | InitializeKMAC256
'|------------------+-------------------------------------------------------
'| Description      | Initialize hasher to use a KMAC256.
'|------------------+-------------------------------------------------------
'| Parameter        | key:                 Key to use for KMAC
'|                  | customizationString: Customization byte array
'|------------------+-------------------------------------------------------
'| Return values    | ./.
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2025-07-14  Created. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | Now, that Windows 11 finally supports KMAC 9 years
'|                  | after it has been standardized by the NIST, we are
'|                  | able to offer it!
'+--------------------------------------------------------------------------
'
Public Sub InitializeKMAC256(ByRef key() As Byte, ByRef customizationString() As Byte)
   InitializeKMAC BCRYPT_KMAC256_ALGORITHM, key, customizationString
End Sub

' Digest methods

'
'+--------------------------------------------------------------------------
'| Method           | DigestPartOfByteArray
'|------------------+-------------------------------------------------------
'| Description      | Hashes part of a byte array.
'|------------------+-------------------------------------------------------
'| Parameter        | sourceByteArray: Source bytes to hash
'|                  | offset         : Offset in byte array
'|                  | count          : Number of bytes to process
'|------------------+-------------------------------------------------------
'| Return values    | ./.
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2022-03-18  Created. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | ./.
'+--------------------------------------------------------------------------
'
Public Sub DigestPartOfByteArray(ByRef sourceByteArray() As Byte, ByVal offset As Long, ByVal count As Long)
   CheckState
   
   Dim startIndex As Long
   startIndex = LBound(sourceByteArray) + offset

   If count > 0 Then
      If startIndex + count - 1 <= UBound(sourceByteArray) Then
         Dim rc As Long

         rc = BCryptHashData(m_HashHandle, sourceByteArray(startIndex), count, 0&)

         If Not IsNTSuccess(rc) Then _
            HandleNTAPIError "hash bytes", "BCryptHashData", rc
      Else
         Err.Raise ERR_NUM_OFFSET_COUNT_EXCEED_UPPER_BOUND, _
                   ERR_STR_CLASS_NAME, _
                   m_MM.FormatMessageWithParameters(ERR_STR_OFFSET_COUNT_EXCEED_UPPER_BOUND, offset, count, UBound(sourceByteArray))
      End If
   End If
End Sub

'
'+--------------------------------------------------------------------------
'| Method           | DigestBytes
'|------------------+-------------------------------------------------------
'| Description      | Hashes a byte array.
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
         HandleNTAPIError "hash bytes", "BCryptHashData", rc
   End If
End Sub

'
'+--------------------------------------------------------------------------
'| Method           | DigestStringAsAnsi
'|------------------+-------------------------------------------------------
'| Description      | Hashes a string and uses Windows ANSI character code
'|                  | conversion from characters to bytes.
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
'|                  | conversion from characters to bytes.
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

'
'+--------------------------------------------------------------------------
'| Method           | DigestFile
'|------------------+-------------------------------------------------------
'| Description      | Hashes the contents of a file.
'|------------------+-------------------------------------------------------
'| Parameter        | filePath: Path of file
'|------------------+-------------------------------------------------------
'| Return values    | ./.
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2022-03-18  Created. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | ./.
'+--------------------------------------------------------------------------
'
Public Sub DigestFile(ByRef filePath As String)
   CheckState
   
   Dim bfis As New BinaryFileInputStream
   bfis.ForFile filePath
   
   Dim buffer(1 To BUFFER_SIZE) As Byte

   Dim rc As Long

   Dim readCount As Long
   Do
      readCount = bfis.Read(buffer)

      If readCount > 0 Then
         rc = BCryptHashData(m_HashHandle, buffer(LBound(buffer)), readCount, 0&)

         If Not IsNTSuccess(rc) Then _
            HandleNTAPIError "hash bytes", "BCryptHashData", rc
      Else
         Exit Do
      End If
   Loop

   bfis.CloseStream

   Set bfis = Nothing
End Sub

' Get method

'
'+--------------------------------------------------------------------------
'| Method           | GetHash
'|------------------+-------------------------------------------------------
'| Description      | Get hash value.
'|------------------+-------------------------------------------------------
'| Parameter        | ./.
'|------------------+-------------------------------------------------------
'| Return values    | Hash value.
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2020-07-18  Created. fhs
'|                  | 2022-02-04  Destroy hash after call to BCryptFinishHash. fhs
'|                  | 2022-03-04  Reverted last change. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | After a call to this method the hash is reset, so it
'|                  | can be used again from the start.
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
      HandleNTAPIError "get hash data", _
                       "BCryptFinishHash", _
                       rc
   End If
End Function

'
' Class methods
'
Private Sub Class_Terminate()
   DestroyProvider   ' Destroy everything on termination to prevent resource leaks
End Sub
