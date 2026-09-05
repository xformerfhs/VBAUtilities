Attribute VB_Name = "AuthenticatedEncryptionFactory"
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
'| Description         | Provide constants, types and a factory function for
'|                     | authenticated encryption.
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
' Public constants
'

'
'+--------------------------------------------------------------------------
'| Enum             | TAuthenticatedEncryptionType
'|------------------+-------------------------------------------------------
'| Description      | Authenticated encryption types.
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2026-09-05  Created. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | Use one of these when calling
'|                  | CreateAuthenticatedEncryption or
'|                  | AuthenticatedEncryptionCng.SetEncryption
'+--------------------------------------------------------------------------
'
Public Enum TAuthenticatedEncryptionType
    aetNotSet
    aetChaCha20Poly1305
    aetAesGcm
    aetAesCcm
End Enum

'
' Public types
'

'
'+--------------------------------------------------------------------------
'| Type             | TAuthenticatedEncryptionResult
'|------------------+-------------------------------------------------------
'| Description      | Data returned from an authenticated encryption.
'|------------------+-------------------------------------------------------
'| Fields           | EncryptedData: The encrypted data.
'|                  | Tag          : The authentication tag.
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2026-09-05  Created. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | ./.
'+--------------------------------------------------------------------------
'
Public Type TAuthenticatedEncryptionResult
   EncryptedData() As Byte
   Tag() As Byte
End Type

'
'+--------------------------------------------------------------------------
'| Type             | TAuthenticatedDecryptionResult
'|------------------+-------------------------------------------------------
'| Description      | Data returned from an authenticated decryption.
'|------------------+-------------------------------------------------------
'| Fields           | ClearData           : The decrypted data or an empty
'|                  |                       byte array if the authentication
'|                  |                       failed.
'|                  | AuthenticationFailed: Flag set to "true", if the
'|                  |                       supplied authentication tag did
'|                  |                       not match the data.
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2026-09-05  Created. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | ./.
'+--------------------------------------------------------------------------
'
Public Type TAuthenticatedDecryptionResult
   ClearData() As Byte
   AuthenticationFailed As Boolean
End Type

'
' Public functions
'

'
'+--------------------------------------------------------------------------
'| Method           | CreateAuthenticatedEncryption
'|------------------+-------------------------------------------------------
'| Description      | Creates an authenticated encryption instance.
'|------------------+-------------------------------------------------------
'| Parameter        | et : Encryption type.
'|                  | key: Key to be used with the instance.
'|------------------+-------------------------------------------------------
'| Return values    | An instance of AuthenticatedEncryptionCng.
'|                  | If the encryption type is invalid an exception
'|                  | is raised.
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2026-09-05  Created. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | The key can be cleared after the creation of the
'|                  | instance.
'|------------------+-------------------------------------------------------
'| Usage            | A typical usage is
'|                  |
'|                  | Dim encryptor As AuthenticatedEncryptionCng
'|                  | Set encryptor = CreateAuthenticatedEncryption(aetChaCha20Poly1305, key)
'|                  | Dim encrypted As TAuthenticatedEncryptionResult
'|                  | encrypted = encryptor.Encrypt(nonce, data, associatedData)
'|                  | ...
'|                  | Dim decrypted As TAuthenticatedDecryptionResult
'|                  | decrypted = encryptor.Decrypt(nonce, encrypted.EncryptedData, associatedData, encrypted.Tag)
'|                  | if decrypted.AuthenticationFailed Then
'|                  |    ' Handle authentication failure
'|                  | End If
'|                  | ' Handle decrypted data
'+--------------------------------------------------------------------------
'
Public Function CreateAuthenticatedEncryption(ByVal et As TAuthenticatedEncryptionType, ByRef key() As Byte) As AuthenticatedEncryptionCng
    Dim instance As New AuthenticatedEncryptionCng
    instance.SetEncryption et, key
    Set CreateAuthenticatedEncryption = instance
End Function
