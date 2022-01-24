VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SecureRandomNumberCng"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
'+-------------------------------------------------------------------------
'|
'| SPDX-FileCopyrightText: 2022 Frank Schwab
'|
'| SPDX-License-Identifier: MIT
'|
'| Copyright 2022, Frank Schwab
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
'| Class               | SecureRandomNumberCng
'|---------------------+---------------------------------------------------
'| Description         | Secure CNG random number generator
'|---------------------+---------------------------------------------------
'| Author              | Frank Schwab
'|---------------------+---------------------------------------------------
'| Version             | 2.0.1
'|---------------------+---------------------------------------------------
'| Changes             | 2020-07-23  Created. fhs
'|                     | 2022-01-22  Removed unnecessary terminating
'|                     |             null character. fhs
'|                     | 2022-01-22  Made 64 bit compatible. fhs
'|                     | 2022-01-23  V2.0.0: Use a hasher to protect against
'|                     |             weaknesses in the CTR_DRBG PRNG of
'|                     |             BCryptGenRandom. fhs
'|                     | 2022-01-24  Slight randomness improvement. fhs
'|---------------------+---------------------------------------------------
'

Option Explicit

' Algorithm ids
Private Const BCRYPT_RNG_ALGORITHM As String = "RNG"

' Failed/Succeeded
Private Const STATUS_SUCCESS As Long = 0


'
' Constants for error messages
'
Private Const STR_ERROR_SOURCE   As String = "SecureRandomNumberCng"
Private Const ERR_BASE As Long = vbObjectError + 64219

' Crypto API error
Private Const STR_ERR_CRYPTO_API As String = "CNG Crypto API function %1 returned code 0x%2: %3"
Private Const ERR_CRYPTO_API As Long = ERR_BASE


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

Private Declare PtrSafe Function BCryptGenRandom Lib "bcrypt" ( _
   ByVal hAlgorithm As LongPtr, _
   ByRef pbBuffer As Byte, _
   ByVal cbBuffer As Long, _
   ByVal dwFlags As Long) As Long

Private Declare PtrSafe Function QueryPerformanceCounter Lib "kernel32" ( _
   ByRef lpPerformanceCount As Byte) As Long

'
' Private constants
'
Private Const BUFFER_SIZE       As Integer = 256
Private Const RANDOM_BLOCK_SIZE As Integer = 16

'
' Instance variables
'
Private m_AlgorithmHandle As LongPtr
Private m_NumberBuffer(1 To BUFFER_SIZE) As Byte
Private m_ActIndex As Integer
Private m_Hasher As HashCng
Private m_RandomBlock(1 To RANDOM_BLOCK_SIZE) As Byte
Private m_BlockCounter(1 To 4) As Byte
Private m_IsBlockCounterInitialized As Boolean

'
' Private methods
'

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
'| Method           | GetHasher
'|------------------+-------------------------------------------------------
'| Description      | Get a hasher
'|------------------+-------------------------------------------------------
'| Parameter        | ./.
'|------------------+-------------------------------------------------------
'| Return values    | A hasher
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2021-01-22  Created. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | ./.
'+--------------------------------------------------------------------------
'
Private Function GetHasher() As HashCng
   If m_Hasher Is Nothing Then
      Set m_Hasher = New HashCng
      m_Hasher.InitializeSHA256Hash
   End If

   Set GetHasher = m_Hasher
End Function

'
'+--------------------------------------------------------------------------
'| Method           | IncrementBlockCounter
'|------------------+-------------------------------------------------------
'| Description      | Increment the block counter
'|------------------+-------------------------------------------------------
'| Parameter        | ./.
'|------------------+-------------------------------------------------------
'| Return values    | ./.
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2019-02-20  Created. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | This method increments the block counter.
'+--------------------------------------------------------------------------
'
Private Sub IncrementBlockCounter()
   Dim i As Long
   For i = UBound(m_BlockCounter) To LBound(m_BlockCounter) Step -1
      Dim actValue As Byte
      
      actValue = m_BlockCounter(i)
      
      If actValue < &HFF Then
         m_BlockCounter(i) = actValue + 1
         Exit For
      Else
         m_BlockCounter(i) = 0
      End If
   Next i
End Sub

'
'+--------------------------------------------------------------------------
'| Method           | CopyBytes
'|------------------+-------------------------------------------------------
'| Description      | Copy bytes from one array to another
'|------------------+-------------------------------------------------------
'| Parameter        | fromArray:      Source array
'|                  | fromStartIndex: First index in source array
'|                  | toArray:        Destination array
'|                  | toStartIndex:   First index in destination array
'|                  | count:          Number of bytes to copy
'|------------------+-------------------------------------------------------
'| Return values    | ./.
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2019-02-20  Created. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | This method increments the block counter.
'+--------------------------------------------------------------------------
'
Private Sub CopyBytes(ByRef fromArray() As Byte, ByVal fromStartIndex As Long, ByRef toArray() As Byte, ByVal toStartIndex As Long, ByVal count As Long)
   Dim toIndex As Long
   toIndex = toStartIndex
   
   Dim doneCount As Long
   
   Dim i As Long
   For i = fromStartIndex To UBound(fromArray)
      If doneCount < count Then
         toArray(toIndex) = fromArray(i)
         
         toIndex = toIndex + 1
         doneCount = doneCount + 1
      Else
         Exit For
      End If
   Next i
End Sub

'
'+--------------------------------------------------------------------------
'| Method           | ClearArray
'|------------------+-------------------------------------------------------
'| Description      | Clear a byte array
'|------------------+-------------------------------------------------------
'| Parameter        | anArray: Byte array to clear
'|------------------+-------------------------------------------------------
'| Return values    | ./.
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2019-02-20  Created. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | ./.
'+--------------------------------------------------------------------------
'
Private Sub ClearArray(ByRef anArray() As Byte)
   Dim i As Long
   
   For i = LBound(anArray) To UBound(anArray)
      anArray(i) = 0
   Next i
End Sub

'
'+--------------------------------------------------------------------------
'| Method           | OpenRNGAlgorithm
'|------------------+-------------------------------------------------------
'| Description      | Open handle of PRNG algorithm
'|------------------+-------------------------------------------------------
'| Parameter        | ./.
'|------------------+-------------------------------------------------------
'| Return values    | Algorithm handle
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2019-02-20  Created. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | ./.
'+--------------------------------------------------------------------------
'
Private Function OpenRNGAlgorithm() As LongPtr
   Dim rc As Long

   Dim result As LongPtr

   rc = BCryptOpenAlgorithmProvider(result, _
                                    StrPtr(BCRYPT_RNG_ALGORITHM), _
                                    0&, _
                                    0&)

   If rc <> STATUS_SUCCESS Then _
      HandleError "BCryptOpenAlgorithmProvider", rc

   OpenRNGAlgorithm = result
End Function

'
'+--------------------------------------------------------------------------
'| Method           | EnsureAlgorithmHandle
'|------------------+-------------------------------------------------------
'| Description      | Ensure that there is an algorithm handle to the PRNG.
'|------------------+-------------------------------------------------------
'| Parameter        | ./.
'|------------------+-------------------------------------------------------
'| Return values    | ./.
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2019-02-20  Created. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | This method set the private instance variable
'|                  | m_AlgorithmHandle.
'+--------------------------------------------------------------------------
'
Private Sub EnsureAlgorithmHandle()
   If m_AlgorithmHandle = 0 Then _
     m_AlgorithmHandle = OpenRNGAlgorithm
End Sub

'
'+--------------------------------------------------------------------------
'| Method           | GetRandomData
'|------------------+-------------------------------------------------------
'| Description      | Copies PRNG data from the generator into an array.
'|------------------+-------------------------------------------------------
'| Parameter        | destinationArray: Array where to put the data.
'|                  | count:            Number of bytes to get.
'|------------------+-------------------------------------------------------
'| Return values    | ./.
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2019-02-20  Created. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | ./.
'+--------------------------------------------------------------------------
'
Private Sub GetRandomData(ByRef destinationArray() As Byte, ByVal count As Long)
   Dim rc As Long

   rc = BCryptGenRandom(m_AlgorithmHandle, _
                        destinationArray(LBound(destinationArray)), _
                        count, _
                        0&)

   If rc <> STATUS_SUCCESS Then _
      HandleError "BCryptGenRandom", rc
End Sub

'
'+--------------------------------------------------------------------------
'| Method           | WhitenBlockCounter
'|------------------+-------------------------------------------------------
'| Description      | Add additional randomness to the block counter.
'|------------------+-------------------------------------------------------
'| Parameter        | ./.
'|------------------+-------------------------------------------------------
'| Return values    | ./.
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2022-01-23  Created. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | ./.
'+--------------------------------------------------------------------------
'
Private Sub WhitenBlockCounter()
      Dim performanceArray(1 To 8) As Byte
      
      QueryPerformanceCounter performanceArray(1)

      Dim fromIndex As Integer
      fromIndex = 1
      
      Dim i As Integer
      For i = LBound(m_BlockCounter) To UBound(m_BlockCounter)
         m_BlockCounter(i) = m_BlockCounter(i) Xor performanceArray(fromIndex)
         fromIndex = fromIndex + 1
      Next i
End Sub

'
'+--------------------------------------------------------------------------
'| Method           | RefreshNumberBuffer
'|------------------+-------------------------------------------------------
'| Description      | Fetch a new buffer of pseudo random numbers.
'|------------------+-------------------------------------------------------
'| Parameter        | ./.
'|------------------+-------------------------------------------------------
'| Return values    | ./.
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2019-02-20  Created. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | This method set the private instance variable
'|                  | m_NumberBuffer. It hashes the result of the
'|                  | CNG PRNG to protected against weaknesses of this
'|                  | implementation.
'+--------------------------------------------------------------------------
'
Private Sub RefreshNumberBuffer()
   Dim hasher As HashCng
   Set hasher = GetHasher

   Dim sourceOffset As Long
   sourceOffset = 2

   Dim maxSourceOffset As Long
   maxSourceOffset = hasher.HashSize - 16

   Dim destinationIndex As Long
   destinationIndex = LBound(m_NumberBuffer)

   EnsureAlgorithmHandle

   If Not m_IsBlockCounterInitialized Then
      GetRandomData m_BlockCounter, 4
      
      WhitenBlockCounter
      
      m_IsBlockCounterInitialized = True
   End If

   Do While destinationIndex <= UBound(m_NumberBuffer)
      GetRandomData m_RandomBlock, RANDOM_BLOCK_SIZE

      hasher.DigestBytes m_RandomBlock
      IncrementBlockCounter
      hasher.DigestBytes m_BlockCounter

      Dim hashResult() As Byte

      hashResult = hasher.GetHash

      CopyBytes hashResult, sourceOffset + LBound(hashResult), m_NumberBuffer, destinationIndex, RANDOM_BLOCK_SIZE
      ClearArray hashResult

      sourceOffset = (sourceOffset + 11) And 15

      destinationIndex = destinationIndex + RANDOM_BLOCK_SIZE
   Loop

   ClearArray m_RandomBlock
End Sub

'
'+--------------------------------------------------------------------------
'| Method           | PutRandomBytesInBuffer
'|------------------+-------------------------------------------------------
'| Description      | Puts the requested no. of pseudo random bytes into
'|                  | the destination buffer.
'|------------------+-------------------------------------------------------
'| Parameter        | aBuffer: Destination buffer
'|                  | requestedNoOfBytes: No. of pseudo random bytes to get
'|------------------+-------------------------------------------------------
'| Return values    | ./.
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2019-02-20  Created. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | ./.
'+--------------------------------------------------------------------------
'
Private Sub PutRandomBytesInBuffer(ByRef aBuffer() As Byte, ByVal requestedNoOfBytes As Integer)
   Dim actDestinationIndex As Integer

   actDestinationIndex = LBound(aBuffer)

   Do While actDestinationIndex <= UBound(aBuffer)
      If m_ActIndex > UBound(m_NumberBuffer) Then
         RefreshNumberBuffer
         m_ActIndex = LBound(m_NumberBuffer)
      End If

      aBuffer(actDestinationIndex) = m_NumberBuffer(m_ActIndex)
      
      m_ActIndex = m_ActIndex + 1
      actDestinationIndex = actDestinationIndex + 1
   Loop
End Sub

'
' Public methods
'

'
'+--------------------------------------------------------------------------
'| Method           | GetRandomByte
'|------------------+-------------------------------------------------------
'| Description      | Gets one random byte
'|------------------+-------------------------------------------------------
'| Parameter        | ./.
'|------------------+-------------------------------------------------------
'| Return values    | ./.
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2019-02-20  Created. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | ./.
'+--------------------------------------------------------------------------
'
Public Function GetRandomByte() As Byte
   GetRandomByte = Me.GetRandomBytes(1)(1)
End Function

'
'+--------------------------------------------------------------------------
'| Method           | GetRandomBytes
'|------------------+-------------------------------------------------------
'| Description      | Gets the requested number of random bytes.
'|------------------+-------------------------------------------------------
'| Parameter        | requestedNoOfBytes: No. of pseudo random bytes to get
'|------------------+-------------------------------------------------------
'| Return values    | ./.
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2019-02-20  Created. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | This method allocates the random byte array.
'+--------------------------------------------------------------------------
'
Public Function GetRandomBytes(ByVal requestedNoOfBytes As Integer) As Byte()
   Dim result() As Byte
   ReDim result(1 To requestedNoOfBytes)

   PutRandomBytesInBuffer result, requestedNoOfBytes

   GetRandomBytes = result
End Function

'
'+--------------------------------------------------------------------------
'| Method           | GetRandomBytesIntoArray
'|------------------+-------------------------------------------------------
'| Description      | Fills an array with random bytes.
'|------------------+-------------------------------------------------------
'| Parameter        | aByteArray: The byte array to fill with pseudo random
'|                  |             bytes.
'|------------------+-------------------------------------------------------
'| Return values    | ./.
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2019-02-20  Created. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | ./.
'+--------------------------------------------------------------------------
'
Public Function GetRandomBytesIntoArray(ByRef aByteArray() As Byte)
   PutRandomBytesInBuffer aByteArray, GetLengthOfByteArray(aByteArray)
End Function

'
' Class methods
'
Private Sub Class_Initialize()
   m_AlgorithmHandle = 0&                    ' No algorithm handle, yet
   m_ActIndex = UBound(m_NumberBuffer) + 1   ' Signal that the buffer has to be filled
End Sub

Private Sub Class_Terminate()
   If m_AlgorithmHandle <> 0 Then _
      BCryptCloseAlgorithmProvider m_AlgorithmHandle, 0&
   
   m_ActIndex = 0
   
   ClearArray m_BlockCounter
   ClearArray m_NumberBuffer
End Sub
