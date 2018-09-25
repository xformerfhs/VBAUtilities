VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SHA1Hash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
'+-------------------------------------------------------------------------
'|
'|                            Copyright (C) 2018
'|                            Frank Schwab
'|
'| Permission is hereby granted, free of charge, to any person obtaining a copy
'| of this software and associated documentation files (the "Software"), to deal
'| in the Software without restriction, including without limitation the rights
'| to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
'| copies of the Software, and to permit persons to whom the Software is
'| furnished to do so, subject to the following conditions:
'|
'| The above copyright notice and this permission notice shall be included in
'| all copies or substantial portions of the Software.
'|
'| THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
'| IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
'| FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
'| AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
'| LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
'| OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
'| THE SOFTWARE.
'|
'|-------------------------------------------------------------------------
'| Class               | SHA1Hash
'|---------------------+---------------------------------------------------
'| Description         | Implements the SHA-1 hash by using the Microsoft
'|                     | CNG API
'|---------------------+---------------------------------------------------
'| Author              | Frank Schwab
'|---------------------+---------------------------------------------------
'| Version             | 4.0.0
'|---------------------+---------------------------------------------------
'| Changes             | 2002-03-15  Created. fhs
'|                     | 2006-03-06  V2: Error handling through exceptions. fhs
'|                     | 2006-03-08  V2.1.0: Get error messages from system. fhs
'|                     | 2010-05-07  V3.0.0: Hash value is a byte array. fhs
'|                     | 2013-03-18  V3.1.0: Simplified. fhs
'|                     | 2014-03-31  V4.0.0: Use CNG. fhs
'|---------------------+---------------------------------------------------
'| Typical usage       | Dim sha1 As New SHA1HashCNG
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

Option Compare Binary
Option Explicit


'
' Constants for error messages
'
Private Const STR_ERROR_SOURCE   As String = "SHA1Hash"
Private Const ERR_BASE As Long = vbObjectError + 1427

' Crypto API error
Private Const STR_ERR_CRYPTO_API As String = "Unable to %1. %2 returned code 0x%3: %4"
Private Const ERR_CRYPTO_API     As Long = ERR_BASE + 0

' Error message when this object has not been used properly
Private Const STR_ERR_INVALID_USAGE As String = "Hash has not been %1"
Private Const ERR_INVALID_USAGE     As Long = ERR_BASE + 1


'
' Private constants
'

' Algorithm ids
Private Const BCRYPT_SHA1_ALGORITHM As String = "SHA1"

' Property names for BCryptGetProperty
Private Const BCRYPT_HASH_LENGTH As String = "HashDigestLength"

' Failed/Succeeded
Private Const STATUS_SUCCESS As Long = 0


'
' API declarations
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
                 ByRef pbHashObject As Any, _
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
                (ByVal phAlgorithm As Long, _
                 ByVal dwFlags As Long) _
                As Long


'
' Instance variables
'
Private m_ContextHandle As Long
Private m_HashHandle    As Long
Private m_HashSize      As Long
Private m_HashValue()   As Byte


'
' Private helper objects
'
Private m_B64 As New Base64Converter
Private m_HC  As New HexConverter
Private m_MM  As New MessageManager
Private m_U8C As New UTF8Converter


'
' Private methods
'

'
'+--------------------------------------------------------------------------
'| Method           | EnsureCryptContext
'|------------------+-------------------------------------------------------
'| Description      | Ensure that there is a valid crypto context
'|------------------+-------------------------------------------------------
'| Parameter        | ./.
'|------------------+-------------------------------------------------------
'| Return values    | ./.
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2014-03-18  Created. fhs
'|                  | 2014-03-31  Use CNG. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | ./.
'+--------------------------------------------------------------------------
'
Private Sub EnsureCryptContext()
   Dim rc As Long

   If m_ContextHandle = 0 Then
      rc = BCryptOpenAlgorithmProvider(m_ContextHandle, _
                                       StrPtr(BCRYPT_SHA1_ALGORITHM), _
                                       0&, _
                                       0&)

      If rc <> STATUS_SUCCESS Then
         Err.Raise ERR_CRYPTO_API, _
                   STR_ERROR_SOURCE, _
                   m_MM.FormatMessageWithParameters(STR_ERR_CRYPTO_API, _
                                                    "open algorithm provider", _
                                                    "BCryptOpenAlgorithmProvider", _
                                                    Hex$(rc), _
                                                    m_MM.GetMessageForNTStatusCode(rc))
      End If
   End If
End Sub

'
'+--------------------------------------------------------------------------
'| Method           | DestroyExistingHash
'|------------------+-------------------------------------------------------
'| Description      | Destroy existing hash object
'|------------------+-------------------------------------------------------
'| Parameter        | ./.
'|------------------+-------------------------------------------------------
'| Return values    | ./.
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2014-03-18  Created. fhs
'|                  | 2014-03-31  Use CNG. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | This method *must* be called after the hash has
'|                  | been finalized and read out.
'+--------------------------------------------------------------------------
'
Private Sub DestroyExistingHash()
   Dim rc As Long

   rc = BCryptDestroyHash(m_HashHandle)

   If rc = STATUS_SUCCESS Then
      m_HashHandle = 0
   Else
      Err.Raise ERR_CRYPTO_API, _
                STR_ERROR_SOURCE, _
                m_MM.FormatMessageWithParameters(STR_ERR_CRYPTO_API, _
                                                 "destroy hash", _
                                                 "BCryptDestroyHash", _
                                                 Hex$(rc), _
                                                 m_MM.GetMessageForNTStatusCode(rc))
   End If
End Sub


'
' Public methods
'

'
'+--------------------------------------------------------------------------
'| Method           | Initialize
'|------------------+-------------------------------------------------------
'| Description      | Initialize hash object
'|------------------+-------------------------------------------------------
'| Parameter        | ./.
'|------------------+-------------------------------------------------------
'| Return values    | ./.
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2002-03-15  Created. fhs
'|                  | 2014-03-18  Simplified. fhs
'|                  | 2014-03-31  Use CNG. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | ./.
'+--------------------------------------------------------------------------
'
Public Sub Initialize()
   Dim rc As Long
   Dim hashSizeSize As Long
   Dim newHashSize As Long

   If m_HashHandle <> 0 Then
      DestroyExistingHash  ' Destroy a hash that may still be there from a former invocation
   Else
      EnsureCryptContext   ' If this is the very first call ensure that there is a crypto context
   End If

   '
   ' Get length of hash
   '
   rc = BCryptGetProperty(m_ContextHandle, _
                          StrPtr(BCRYPT_HASH_LENGTH), _
                          newHashSize, _
                          LenB(newHashSize), _
                          hashSizeSize, _
                          0&)

   If rc = STATUS_SUCCESS Then
      '
      ' Allocate a new byte array to store the hash value
      ' if necessary. This is always done on the very first call.
      '
      If newHashSize <> m_HashSize Then
         m_HashSize = newHashSize
         ReDim m_HashValue(1 To newHashSize)
      End If
   
      '
      ' Create the hash object
      '
      rc = BCryptCreateHash(m_ContextHandle, _
                            m_HashHandle, _
                            ByVal 0&, _
                            0&, _
                            0&, _
                            0&, _
                            0&)

      If rc <> STATUS_SUCCESS Then
         Err.Raise ERR_CRYPTO_API, _
                   STR_ERROR_SOURCE, _
                   m_MM.FormatMessageWithParameters(STR_ERR_CRYPTO_API, _
                                                    "create hash", _
                                                    "BCryptCreateHash", _
                                                    Hex$(rc), _
                                                    m_MM.GetMessageForNTStatusCode(rc))
      
      End If
   Else
      Err.Raise ERR_CRYPTO_API, _
                STR_ERROR_SOURCE, _
                m_MM.FormatMessageWithParameters(STR_ERR_CRYPTO_API, _
                                                 "get hash size", _
                                                 "BCryptGetProperty", _
                                                 Hex$(rc), _
                                                 m_MM.GetMessageForNTStatusCode(rc))
   End If
End Sub

'
'+--------------------------------------------------------------------------
'| Method           | DigestBytes
'|------------------+-------------------------------------------------------
'| Description      | Digest a byte array
'|------------------+-------------------------------------------------------
'| Parameter        | sourceData: Data to be hashed
'|------------------+-------------------------------------------------------
'| Return values    | ./.
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2010-05-07  Created. fhs
'|                  | 2014-03-18  Simplified. fhs
'|                  | 2014-03-31  Use CNG. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | ./.
'+--------------------------------------------------------------------------
'
Public Sub DigestBytes(ByRef sourceData() As Byte)
   Dim rc As Long
   
   If m_HashHandle = 0 Then
      Me.Initialize
   End If

   rc = BCryptHashData(m_HashHandle, _
                       sourceData(LBound(sourceData)), _
                       UBound(sourceData) - LBound(sourceData) + 1, _
                       0&)

   If rc <> STATUS_SUCCESS Then
      Err.Raise ERR_CRYPTO_API, _
                STR_ERROR_SOURCE, _
                m_MM.FormatMessageWithParameters(STR_ERR_CRYPTO_API, _
                                                 "hash data", _
                                                 "BCryptHashData", _
                                                 Hex$(rc), _
                                                 m_MM.GetMessageForNTStatusCode(rc))
   End If
End Sub

'
'+--------------------------------------------------------------------------
'| Method           | DigestString
'|------------------+-------------------------------------------------------
'| Description      | Digest a string
'|------------------+-------------------------------------------------------
'| Parameter        | sourceString: String to be digested
'|------------------+-------------------------------------------------------
'| Return values    | ./.
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2002-03-15  Created. fhs
'|                  | 2010-05-06  Convert string to UTF-8 representation
'|                  |             before digesting. fhs
'|                  | 2014-03-18  Simplified, correct handling of empty
'|                  |             Strings. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | The string is converted to UTF-8 representation and
'|                  | this UTF-8 byte array is digested.
'|                  |
'|                  | An empty string is just ignored. That is the correct
'|                  | way to handle it.
'+--------------------------------------------------------------------------
'
Public Sub DigestString(ByRef sourceString As String)
   If Len(sourceString) > 0 Then
      Me.DigestBytes m_U8C.FromVBToUTF8(sourceString)
   End If
End Sub

'
'+--------------------------------------------------------------------------
'| Method           | Finish
'|------------------+-------------------------------------------------------
'| Description      | Finialize hash and get hash value
'|------------------+-------------------------------------------------------
'| Parameter        | ./.
'|------------------+-------------------------------------------------------
'| Return values    | Hash value in m_Hash
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2002-03-15  Created. fhs
'|                  | 2014-03-31  Use CNG. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | ./.
'+--------------------------------------------------------------------------
'
Public Sub Finish()
   Dim rc As Long

   If m_HashHandle = 0 Then
      Err.Raise ERR_INVALID_USAGE, _
                STR_ERROR_SOURCE, _
                m_MM.FormatMessageWithParameters(STR_ERR_INVALID_USAGE, "initialized")
   Else
      rc = BCryptFinishHash(m_HashHandle, _
                            m_HashValue(LBound(m_HashValue)), _
                            UBound(m_HashValue) - LBound(m_HashValue) + 1, _
                            0&)

      If rc = STATUS_SUCCESS Then
         DestroyExistingHash
      Else
         Err.Raise ERR_CRYPTO_API, _
                   STR_ERROR_SOURCE, _
                   m_MM.FormatMessageWithParameters(STR_ERR_CRYPTO_API, _
                                                    "finish hash", _
                                                    "BCryptFinishHash", _
                                                    Hex$(rc), _
                                                    m_MM.GetMessageForNTStatusCode(rc))
      End If
   End If
End Sub

'
' Public properties
'

'
'+--------------------------------------------------------------------------
'| Property         | HashValue
'|------------------+-------------------------------------------------------
'| Description      | Hashwert auslesen
'|------------------+-------------------------------------------------------
'| Parameter        | ./.
'|------------------+-------------------------------------------------------
'| Return values    | Hash value as byte array
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2002-03-15  Created. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | ./.
'+--------------------------------------------------------------------------
'
Public Property Get HashValue() As Byte()
   If m_HashHandle <> 0 Then
      Me.Finish
   End If

   If m_HashSize <> 0 Then
      HashValue = m_HashValue
   Else
      Err.Raise ERR_INVALID_USAGE, _
                STR_ERROR_SOURCE, _
                m_MM.FormatMessageWithParameters(STR_ERR_INVALID_USAGE, "initialized")
   End If
End Property

'
'+--------------------------------------------------------------------------
'| Property         | HashValueAsHexString
'|------------------+-------------------------------------------------------
'| Description      | Hash value as hex string
'|------------------+-------------------------------------------------------
'| Parameter        | ./.
'|------------------+-------------------------------------------------------
'| Return values    | Hash value as hex string
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2010-05-06  Created. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | ./.
'+--------------------------------------------------------------------------
'
Public Property Get HashValueAsHexString() As String
   HashValueAsHexString = m_HC.Encode(Me.HashValue)
End Property

'
'+--------------------------------------------------------------------------
'| Property         | HashValueAsBase64String
'|------------------+-------------------------------------------------------
'| Description      | Hash value as base64 string
'|------------------+-------------------------------------------------------
'| Parameter        | ./.
'|------------------+-------------------------------------------------------
'| Return values    | Hash value as base64 string
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2000-03-18  Created. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | ./.
'+--------------------------------------------------------------------------
'
Public Property Get HashValueAsBase64String() As String
   HashValueAsBase64String = m_B64.Encode(Me.HashValue)
End Property


'
' Class methods
'

'
'+--------------------------------------------------------------------------
'| Method           | Class_Terminate
'|------------------+-------------------------------------------------------
'| Description      | Destructor
'|------------------+-------------------------------------------------------
'| Parameter        | ./.
'|------------------+-------------------------------------------------------
'| Return values    | ./.
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2002-03-15  Created. fhs
'|                  | 2013-03-18  Simplified. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | ./.
'+--------------------------------------------------------------------------
'
Private Sub Class_Terminate()
   Dim rc As Long

   If m_HashHandle <> 0 Then
      rc = BCryptDestroyHash(m_HashHandle)

      If rc <> STATUS_SUCCESS Then
         Err.Raise ERR_CRYPTO_API, _
                   STR_ERROR_SOURCE, _
                   m_MM.FormatMessageWithParameters(STR_ERR_CRYPTO_API, _
                                                    "destroy hash", _
                                                    "BCryptDestroyHash", _
                                                    Hex$(rc), _
                                                    m_MM.GetMessageForNTStatusCode(rc))
      End If
   End If

   If m_ContextHandle <> 0 Then
      rc = BCryptCloseAlgorithmProvider(m_ContextHandle, 0&)

      If rc <> STATUS_SUCCESS Then
         Err.Raise ERR_CRYPTO_API, _
                   STR_ERROR_SOURCE, _
                   m_MM.FormatMessageWithParameters(STR_ERR_CRYPTO_API, _
                                                    "close algorithm provider", _
                                                    "BCryptCloseAlgorithmProvider", _
                                                    Hex$(rc), _
                                                    m_MM.GetMessageForNTStatusCode(rc))
      End If
   End If
End Sub
