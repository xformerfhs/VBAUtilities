VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AuthenticatedEncryptionCng"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
'+-------------------------------------------------------------------------
'|
'| SPDX-FileCopyrightText: 2026 Frank Schwab
'|
'| SPDX-License-Identifier: MIT
'|
'| Copyright 2026, Frank Schwab
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
'| Class               | AuthenticatedEncryptionCng
'|---------------------+---------------------------------------------------
'| Description         | Provide authenticated encryption
'|---------------------+---------------------------------------------------
'| Author              | Frank Schwab
'|---------------------+---------------------------------------------------
'| Version             | 1.0.0
'|---------------------+---------------------------------------------------
'| Changes             | 2026-09-05  Created. fhs
'|---------------------+---------------------------------------------------
'

Option Explicit

'
' ======== Private entities ========
'

'
' Private constants for error messages
'
Private Const ERR_STR_CLASS_NAME As String = "AuthenticatedEncryptionCng"
Private Const ERR_NUM_BASE As Long = vbObjectError + 7314

' Crypto API error
Private Const ERR_STR_BCRYPT_API As String = "Error in function %1: '%2' returned code 0x%3: %4"
Private Const ERR_NUM_BCRYPT_API As Long = ERR_NUM_BASE

' Invalid encryption type
Private Const ERR_STR_WRONG_TYPE As String = "Invalid encryption type: %1"
Private Const ERR_NUM_WRONG_TYPE As Long = ERR_NUM_BASE + 1

' Type not set
Private Const ERR_STR_TYPE_NOT_SET As String = "Authenticated encryption type has not been set"
Private Const ERR_NUM_TYPE_NOT_SET As Long = ERR_NUM_BASE + 2

'
' Bcrypt entities
'

'
' Predefined algorithm handles
'
Private Const BCRYPT_AES_CCM_ALG_HANDLE As Long = &H1D1
Private Const BCRYPT_AES_GCM_ALG_HANDLE As Long = &H1E1
Private Const BCRYPT_CHACHA20_POLY1305_ALG_HANDLE As Long = &H3A1

'
' API constants
'
Private Const BCRYPT_OBJECT_LENGTH As String = "ObjectLength"
Private Const BCRYPT_AUTH_TAG_LENGTH As String = "AuthTagLength"

'
' API structures
'
Private Type BCRYPT_AUTH_TAG_LENGTHS_STRUCT
   dwMinLength As Long
   dwMaxLength As Long
   dwIncrement As Long
End Type

Private Type BCRYPT_AUTHENTICATED_CIPHER_MODE_INFO
   cbSize As Long
   dwInfoVersion As Long
   pbNonce As LongPtr
   cbNonce As Long
   pbAuthData As LongPtr
   cbAuthData As Long
   pbTag As LongPtr
   cbTag As Long
   pbMacContext As LongPtr
   cbMacContext As Long
   cbAAD As Long
   cbData As LongLong
   dwFlags As Long
End Type

'
' API declarations
'
Private Declare PtrSafe Function BCryptGetProperty Lib "bcrypt.dll" ( _
   ByVal hObject As LongPtr, _
   ByVal pszProperty As LongPtr, _
   ByRef pbOutput As Any, _
   ByVal cbOutput As Long, _
   ByRef pcbResult As Long, _
   ByVal dwFlags As Long) _
   As Long

Private Declare PtrSafe Function BCryptGenerateSymmetricKey Lib "bcrypt.dll" ( _
   ByVal hAlgorithm As LongPtr, _
   ByRef hKey As LongPtr, _
   ByRef pbKeyObject As Any, _
   ByVal cbKeyObject As Long, _
   ByRef pbSecret As Any, _
   ByVal cbSecret As Long, _
   ByVal dwFlags As Long) _
   As Long

Private Declare PtrSafe Function BCryptEncrypt Lib "bcrypt.dll" ( _
   ByVal hKey As LongPtr, _
   ByRef pbInput As Any, _
   ByVal cbInput As Long, _
   ByRef pPaddingInfo As Any, _
   ByRef pbIV As Any, _
   ByVal cbIV As Long, _
   ByRef pbOutput As Any, _
   ByVal cbOutput As Long, _
   ByRef pcbResult As Long, _
   ByVal dwFlags As Long) _
   As Long

Private Declare PtrSafe Function BCryptDecrypt Lib "bcrypt.dll" ( _
   ByVal hKey As LongPtr, _
   ByRef pbInput As Any, _
   ByVal cbInput As Long, _
   ByRef pPaddingInfo As Any, _
   ByRef pbIV As Any, _
   ByVal cbIV As Long, _
   ByRef pbOutput As Any, _
   ByVal cbOutput As Long, _
   ByRef pcbResult As Long, _
   ByVal dwFlags As Long) _
   As Long

Private Declare PtrSafe Function BCryptDestroyKey Lib "bcrypt.dll" ( _
   ByVal hKey As LongPtr) _
   As Long

Private Declare PtrSafe Sub RtlZeroMemory Lib "Kernel32.dll" ( _
   ByRef Destination As Any, _
   ByVal Length As LongPtr)

'
' API status codes
'
Private Const STATUS_AUTH_TAG_MISMATCH As Long = &HC000A002

'
' Instance variables
'
Private m_MM As New MessageManager

Private m_AlgHandle As Long
Private m_KeyHandle As LongPtr
Private m_KeyObject() As Byte
Private m_TagLength As Long

'
' ======== Private functions ========
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
'| Method           | ByteArrayLength
'|------------------+-------------------------------------------------------
'| Description      | Get the length of a byte array.
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
Private Function ByteArrayLength(ByRef aByteArray() As Byte) As Long
   On Error Resume Next

   ByteArrayLength = UBound(aByteArray) - LBound(aByteArray) + 1
End Function

'
'+--------------------------------------------------------------------------
'| Method           | ClearByteArray
'|------------------+-------------------------------------------------------
'| Description      | Clear a byte array.
'|------------------+-------------------------------------------------------
'| Parameter        | aByteArray: Byte array to clear
'|------------------+-------------------------------------------------------
'| Return values    | ./.
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2026-09-04  Created. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | ./.
'+--------------------------------------------------------------------------
'
Private Sub ClearByteArray(ByRef aByteArray() As Byte)
   RtlZeroMemory aByteArray(LBound(aByteArray)), ByteArrayLength(aByteArray)
End Sub

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
'| Method           | AuthenticatedEncryptionToAlgHandle
'|------------------+-------------------------------------------------------
'| Description      | Convert an encryption type into an algorithm
'|                  | pseudo-handle.
'|------------------+-------------------------------------------------------
'| Parameter        | et: Encryption type
'|------------------+-------------------------------------------------------
'| Return values    | Algorithm pseudo-handle or 0, when the encryption
'|                  | type is invalid.
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2026-09-04  Created. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | ./.
'+--------------------------------------------------------------------------
'
Private Function AuthenticatedEncryptionToAlgHandle(ByVal et As TAuthenticatedEncryptionType) As Long
   AuthenticatedEncryptionToAlgHandle = 0

   Select Case et
   Case aetChaCha20Poly1305:
      AuthenticatedEncryptionToAlgHandle = BCRYPT_CHACHA20_POLY1305_ALG_HANDLE

   Case aetAesGcm:
      AuthenticatedEncryptionToAlgHandle = BCRYPT_AES_GCM_ALG_HANDLE

   Case aetAesCcm:
      AuthenticatedEncryptionToAlgHandle = BCRYPT_AES_CCM_ALG_HANDLE
   End Select
End Function

'
'+--------------------------------------------------------------------------
'| Method           | SetAuthenticatedCipherModeInfo
'|------------------+-------------------------------------------------------
'| Description      | Set the fields of a BCRYPT_AUTHENTICATED_CIPHER_MODE_INFO
'|                  | structure.
'|------------------+-------------------------------------------------------
'| Parameter        | info          : BCRYPT_AUTHENTICATED_CIPHER_MODE_INFO
'|                  |                 structure
'|                  | nonce         : Nonce
'|                  | associatedData: Associated data
'|                  | tag           : Tag
'|------------------+-------------------------------------------------------
'| Return values    | ./.
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2026-09-04  Created. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | ./.
'+--------------------------------------------------------------------------
'
Private Sub SetAuthenticatedCipherModeInfo( _
      ByRef info As BCRYPT_AUTHENTICATED_CIPHER_MODE_INFO, _
      ByRef nonce() As Byte, _
      ByRef associatedData() As Byte, _
      ByRef Tag() As Byte)
   info.cbSize = LenB(info)
   info.dwInfoVersion = 1
   info.pbNonce = VarPtr(nonce(LBound(nonce)))
   info.cbNonce = ByteArrayLength(nonce)
   If ByteArrayLength(associatedData) <> 0 Then
      info.pbAuthData = VarPtr(associatedData(LBound(associatedData)))
      info.cbAuthData = ByteArrayLength(associatedData)
   End If
   info.pbTag = VarPtr(Tag(LBound(Tag)))
   info.cbTag = ByteArrayLength(Tag)
End Sub

'
'+--------------------------------------------------------------------------
'| Method           | BuildOperationData
'|------------------+-------------------------------------------------------
'| Description      | Build the key data needed for encryption and decryption
'|                  | operations.
'|------------------+-------------------------------------------------------
'| Parameter        | key      : Key
'|                  | hKey     : Key handle
'|                  | keyObject: Key object
'|------------------+-------------------------------------------------------
'| Return values    | ./.
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2026-09-04  Created. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | ./.
'+--------------------------------------------------------------------------
'
Private Sub BuildOperationData( _
      ByRef key() As Byte, _
      ByRef hKey As LongPtr, _
      ByRef keyObject() As Byte)
   '
   ' 1. Clear possible former key data
   '
   ClearKeyData
   
   '
   ' 2. Get the length of the key object.
   '
   Dim rc As Long
   Dim count As Long
   Dim keyObjectLength As Long
   rc = BCryptGetProperty(m_AlgHandle, _
                          StrPtr(BCRYPT_OBJECT_LENGTH), _
                          keyObjectLength, _
                          LenB(keyObjectLength), _
                          count, _
                          0)
   If Not IsNTSuccess(rc) Then
      HandleNTAPIError "BuildOperationData", "BCryptGetProperty::KeyLength", rc
   End If

   '
   ' 3. Generate the key object from the key.
   '
   ReDim keyObject(1 To keyObjectLength)

   rc = BCryptGenerateSymmetricKey(m_AlgHandle, hKey, keyObject(LBound(key)), keyObjectLength, key(LBound(key)), ByteArrayLength(key), 0)
   If Not IsNTSuccess(rc) Then
      HandleNTAPIError "BuildOperationData", "BCryptGenerateSymmetricKey", rc
   End If
   
   '
   ' 4. Get tag length.
   '
   Dim tagLength As BCRYPT_AUTH_TAG_LENGTHS_STRUCT
   rc = BCryptGetProperty(m_AlgHandle, _
                          StrPtr(BCRYPT_AUTH_TAG_LENGTH), _
                          tagLength, _
                          LenB(tagLength), _
                          count, _
                          0)
   If Not IsNTSuccess(rc) Then
      HandleNTAPIError "BuildOperationData", "BCryptGetProperty::TagLength", rc
   End If
   
   m_TagLength = tagLength.dwMaxLength
End Sub

'
'+--------------------------------------------------------------------------
'| Method           | Cleanup
'|------------------+-------------------------------------------------------
'| Description      | Clean up encryption data.
'|------------------+-------------------------------------------------------
'| Parameter        | hKey          : Key handle
'|                  | keyObject     : Key object
'|                  | cipherModeInfo: BCRYPT_AUTHENTICATED_CIPHER_MODE_INFO
'|                  |                 structure
'|------------------+-------------------------------------------------------
'| Return values    | ./.
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2026-09-04  Created. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | ./.
'+--------------------------------------------------------------------------
'
Private Sub Cleanup( _
      ByVal hKey As LongPtr, _
      ByRef keyObject() As Byte, _
      ByRef cipherModeInfo As BCRYPT_AUTHENTICATED_CIPHER_MODE_INFO)
   BCryptDestroyKey hKey

   ClearByteArray keyObject
   RtlZeroMemory cipherModeInfo, cipherModeInfo.cbSize
End Sub

'
'+--------------------------------------------------------------------------
'| Method           | ClearKeyData
'|------------------+-------------------------------------------------------
'| Description      | Clear the key data.
'|------------------+-------------------------------------------------------
'| Parameter        | ./.
'|------------------+-------------------------------------------------------
'| Return values    | ./.
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2026-09-04  Created. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | ./.
'+--------------------------------------------------------------------------
'
Private Sub ClearKeyData()
   If m_KeyHandle <> 0 Then
      BCryptDestroyKey m_KeyHandle
   End If
   
   If ByteArrayLength(m_KeyObject) <> 0 Then
      ClearByteArray m_KeyObject
   End If
End Sub

'
' ======== Public properties ========
'

'
'+--------------------------------------------------------------------------
'| Property         | EncryptionType
'|------------------+-------------------------------------------------------
'| Description      | Get the encryption type used by this instance.
'|------------------+-------------------------------------------------------
'| Parameter        | ./.
'|------------------+-------------------------------------------------------
'| Return values    | The encryption type used by this instance.
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2026-09-05  Created. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | ./.
'+--------------------------------------------------------------------------
'
Public Property Get EncryptionType() As TAuthenticatedEncryptionType
   Select Case m_AlgHandle
   Case BCRYPT_CHACHA20_POLY1305_ALG_HANDLE:
      EncryptionType = aetChaCha20Poly1305

   Case BCRYPT_AES_GCM_ALG_HANDLE:
      EncryptionType = aetAesGcm

   Case BCRYPT_AES_CCM_ALG_HANDLE:
      EncryptionType = aetAesCcm
      
   Case Else
      EncryptionType = aetNotSet
   End Select
End Property

'
'+--------------------------------------------------------------------------
'| Property         | EncryptionTypeString
'|------------------+-------------------------------------------------------
'| Description      | Get the encryption type used by this instance as
'|                  | a string.
'|------------------+-------------------------------------------------------
'| Parameter        | ./.
'|------------------+-------------------------------------------------------
'| Return values    | The encryption type used by this instance as a string.
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2026-09-05  Created. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | ./.
'+--------------------------------------------------------------------------
'
Public Property Get EncryptionTypeString() As String
   Select Case m_AlgHandle
   Case BCRYPT_CHACHA20_POLY1305_ALG_HANDLE:
      EncryptionTypeString = "ChaCha20-Poly1305"

   Case BCRYPT_AES_GCM_ALG_HANDLE:
      EncryptionTypeString = "AES-GCM"

   Case BCRYPT_AES_CCM_ALG_HANDLE:
      EncryptionTypeString = "AES-CCM"
      
   Case Else
      EncryptionTypeString = "*NotSet*"
   End Select
End Property

'
' ======== Public functions ========
'

'
'+--------------------------------------------------------------------------
'| Method           | SetEncryption
'|------------------+-------------------------------------------------------
'| Description      | Set the encryption type.
'|------------------+-------------------------------------------------------
'| Parameter        | encryptionType: The encryption type.
'|------------------+-------------------------------------------------------
'| Return values    | The decrypted bytes. If authentication fails the
'|                  | result is an empty array and the property
'|                  | AuthenticationFailed contains the value "true".
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2026-09-05  Created. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | ./.
'+--------------------------------------------------------------------------
'
Public Sub SetEncryption(ByVal et As TAuthenticatedEncryptionType, ByRef key() As Byte)
   m_AlgHandle = AuthenticatedEncryptionToAlgHandle(et)
   
   If m_AlgHandle = 0 Then
      Err.Raise ERR_NUM_WRONG_TYPE, _
                ERR_STR_CLASS_NAME, _
                m_MM.FormatMessageWithParameters(ERR_STR_WRONG_TYPE, Format$(et))
   End If

   BuildOperationData key, m_KeyHandle, m_KeyObject
End Sub

'
'+--------------------------------------------------------------------------
'| Method           | Encrypt
'|------------------+-------------------------------------------------------
'| Description      | Encrypt data with authenticated encryption.
'|------------------+-------------------------------------------------------
'| Parameter        | nonce         : The nonce for the encryption
'|                  | clearData     : The clear data
'|                  | associatedData: Associated data that are used for the
'|                  |                 tag calculation but that are not
'|                  |                 encrypted (may be an empty array)
'|------------------+-------------------------------------------------------
'| Return values    | A TAuthenticatedEncryptionResult with "EncryptedData"
'|                  | containing the encrypted data and "Tag" containing
'|                  | the tag.
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2026-09-04  Created. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | ./.
'+--------------------------------------------------------------------------
'
Public Function Encrypt( _
      ByRef p_Nonce() As Byte, _
      ByRef p_ClearData() As Byte, _
      ByRef p_AssociatedData() As Byte) As TAuthenticatedEncryptionResult
   Dim result As TAuthenticatedEncryptionResult

   '
   ' 1. Check if instance has been initialized.
   '
   If m_AlgHandle = 0 Then
      Err.Raise ERR_NUM_TYPE_NOT_SET, _
                ERR_STR_CLASS_NAME, _
                ERR_STR_TYPE_NOT_SET
   End If

   '
   ' 2. Build cipher mode info.
   '
   ReDim result.Tag(1 To m_TagLength)
   Dim cipherModeInfo As BCRYPT_AUTHENTICATED_CIPHER_MODE_INFO
   SetAuthenticatedCipherModeInfo cipherModeInfo, p_Nonce, p_AssociatedData, result.Tag

   '
   ' 3. Encrypt the clear data.
   '
   Dim rc As Long
   Dim count As Long
   ReDim result.EncryptedData(1 To ByteArrayLength(p_ClearData))
   rc = BCryptEncrypt( _
      m_KeyHandle, _
      p_ClearData(LBound(p_ClearData)), _
      ByteArrayLength(p_ClearData), _
      cipherModeInfo, _
      ByVal 0, _
      0, _
      result.EncryptedData(LBound(result.EncryptedData)), _
      ByteArrayLength(result.EncryptedData), _
      count, _
      0)
   If Not IsNTSuccess(rc) Then
      RtlZeroMemory cipherModeInfo, cipherModeInfo.cbSize
      HandleNTAPIError "AuthenticatedEncryption", "BCryptEncrypt", rc
   End If

   Encrypt = result

   '
   ' 4. Clean up.
   '
   RtlZeroMemory cipherModeInfo, cipherModeInfo.cbSize
End Function

'
'+--------------------------------------------------------------------------
'| Method           | Decrypt
'|------------------+-------------------------------------------------------
'| Description      | Decrypt data with authenticated encryption.
'|------------------+-------------------------------------------------------
'| Parameter        | nonce         : The nonce for the encryption
'|                  | encryptedData : The encrypted data
'|                  | associatedData: Associated data that are used for the
'|                  |                 tag calculation but that are not
'|                  |                 encrypted (may be an empty array)
'|                  | tag           : The tag
'|------------------+-------------------------------------------------------
'| Return values    | A TAuthenticatedDecryptionResult with "ClearData"
'|                  | set to the decrypted bytes and "AuthenticationFailed"
'|                  | set to "true" if the tag does not match the data.
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2026-09-04  Created. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | ./.
'+--------------------------------------------------------------------------
'
Public Function Decrypt( _
      ByRef p_Nonce() As Byte, _
      ByRef p_EncryptedData() As Byte, _
      ByRef p_AssociatedData() As Byte, _
      ByRef p_Tag() As Byte) As TAuthenticatedDecryptionResult
   Dim result As TAuthenticatedDecryptionResult

   '
   ' 1. Check if instance has been initialized.
   '
   If m_AlgHandle = 0 Then
      Err.Raise ERR_NUM_TYPE_NOT_SET, _
                ERR_STR_CLASS_NAME, _
                ERR_STR_TYPE_NOT_SET
   End If

   '
   ' 2. Build the structure needed for authenticated encryption.
   '
   Dim cipherModeInfo As BCRYPT_AUTHENTICATED_CIPHER_MODE_INFO
   SetAuthenticatedCipherModeInfo cipherModeInfo, p_Nonce, p_AssociatedData, p_Tag

   '
   ' 3. Decrypt the encrypted data.
   '
   Dim rc As Long
   Dim count As Long
   ReDim result.ClearData(1 To ByteArrayLength(p_EncryptedData))
   rc = BCryptDecrypt( _
      m_KeyHandle, _
      p_EncryptedData(LBound(p_EncryptedData)), _
      ByteArrayLength(p_EncryptedData), _
      cipherModeInfo, _
      ByVal 0, _
      0, _
      result.ClearData(LBound(result.ClearData)), _
      ByteArrayLength(result.ClearData), _
      count, _
      0)
   If Not IsNTSuccess(rc) Then
      ClearByteArray result.ClearData
      Dim emptyArray() As Byte
      result.ClearData = emptyArray

      If rc = STATUS_AUTH_TAG_MISMATCH Then
         result.AuthenticationFailed = True
      Else
         RtlZeroMemory cipherModeInfo, cipherModeInfo.cbSize
         HandleNTAPIError "AuthenticatedDecryption", "BCryptDecrypt", rc
      End If
   End If

   Decrypt = result

   '
   ' 4. Clean up.
   '
   RtlZeroMemory cipherModeInfo, cipherModeInfo.cbSize
End Function

'
' Class functions
'

Private Sub Class_Terminate()
   ClearKeyData
End Sub
