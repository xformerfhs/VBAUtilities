VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "InsertionSort"
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
'| Class               | InsertionSort
'|---------------------+---------------------------------------------------
'| Description         | Insertion sort implementation
'|---------------------+---------------------------------------------------
'| Author              | Frank Schwab
'|---------------------+---------------------------------------------------
'| Version             | 1.1.0
'|---------------------+---------------------------------------------------
'| Changes             | 2020-09-25  Created. fhs
'|                     | 2020-09-26  Parameters are checked. fhs
'|---------------------+---------------------------------------------------
'| Remarks             | Insertion sort is only efficient for small arrays.
'|                     | use QuickSort (not PureQuickSort) for efficient
'|                     | sorting of arrays of any size.
'|---------------------+---------------------------------------------------
'

Option Explicit

'
' Constants for errors
'
Private Const ERR_STR_CLASS_NAME As String = "InsertionSort"

Private Const ERR_NUM_START As Long = vbObjectError + 34460

Private Const ERR_NUM_NO_ARRAY As Long = ERR_NUM_START
Private Const ERR_STR_NO_ARRAY As String = "Supplied parameter is not an array"

Private Const ERR_NUM_INVALID_BOUNDARY As Long = ERR_NUM_START + 1
Private Const ERR_STR_INVALID_BOUNDARY_LEFT As String = "Invalid "
Private Const ERR_STR_INVALID_BOUNDARY_RIGHT As String = " boundary"

'
' Private methods
'

'
'+--------------------------------------------------------------------------
'| Method           | GetArrayLength
'|------------------+-------------------------------------------------------
'| Description      | Get the length of an array
'|------------------+-------------------------------------------------------
'| Parameter        | aByteArray: Array to get the length for
'|------------------+-------------------------------------------------------
'| Return values    | Length of array
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2020-09-26  Created. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | When one tries to calculate the length of an
'|                  | uninitialized array an error is raised.
'|                  | This function catches this error and returns a
'|                  | length of 0, instead.
'+--------------------------------------------------------------------------
'
Private Function GetArrayLength(ByRef anArray As Variant) As Long
   On Error Resume Next

   GetArrayLength = UBound(anArray) - LBound(anArray) + 1

   ' If the array is empty there was an error so we set
   ' the return value accordingly

   If Err.Number <> 0 Then _
      GetArrayLength = 0
End Function


'
'+--------------------------------------------------------------------------
'| Method           | InsertionSortWithBoundaries
'|------------------+-------------------------------------------------------
'| Description      | Sort an array of any data type with insertion sort
'|------------------+-------------------------------------------------------
'| Parameter        | arrayToSort: Array to sort
'|                  | idxFrom    : Start index into the array
'|                  | idxTo      : End   index into the array
'|------------------+-------------------------------------------------------
'| Return values    | ./.
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2018-09-25  Created. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | The array that is supplied as the parameter is
'|                  | changed so that its elements from idxFrom to idxTo
'|                  | are sorted.
'+--------------------------------------------------------------------------
'
Private Sub InsertionSortWithBoundaries(ByRef arrayToSort As Variant, ByVal idxFrom As Long, ByVal idxTo As Long)
   Dim idxLeft As Long
   Dim idxLeftForTest As Long
   Dim idxRight As Long
   Dim valueToInsert As Variant
   Dim compareValue As Variant

   For idxRight = idxFrom + 1 To idxTo
      valueToInsert = arrayToSort(idxRight)
      idxLeft = idxRight

      Do
         idxLeftForTest = idxLeft - 1

         If idxLeftForTest >= idxFrom Then
            compareValue = arrayToSort(idxLeftForTest)
            If compareValue > valueToInsert Then
               arrayToSort(idxLeft) = compareValue
               idxLeft = idxLeftForTest
            Else
               Exit Do
            End If
         Else
            Exit Do
         End If
      Loop

      arrayToSort(idxLeft) = valueToInsert
   Next idxRight
End Sub


'
' Public methods
'

'
'+--------------------------------------------------------------------------
'| Method           | SortPart
'|------------------+-------------------------------------------------------
'| Description      | Sort a part of an array of any data type
'|                  | with insertion sort
'|------------------+-------------------------------------------------------
'| Parameter        | arrayToSort: Array to sort
'|                  | idxFrom    : Start index into the array
'|                  | idxTo      : End   index into the array
'|------------------+-------------------------------------------------------
'| Return values    | ./.
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2018-09-25  Created. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | The array that is supplied as the parameter is
'|                  | changed so that its elements from idxFrom to idxTo
'|                  | are sorted.
'+--------------------------------------------------------------------------
'
Public Sub SortPart(ByRef arrayToSort As Variant, ByVal idxFrom As Long, ByVal idxTo As Long)
   If IsArray(arrayToSort) Then
      Dim arrayLength As Long
      arrayLength = GetArrayLength(arrayToSort)

      If arrayLength > 1 Then
         If idxFrom < LBound(arrayToSort) Then
            Err.Raise ERR_NUM_INVALID_BOUNDARY, _
                      ERR_STR_CLASS_NAME, _
                      ERR_STR_INVALID_BOUNDARY_LEFT & "left" & ERR_STR_INVALID_BOUNDARY_RIGHT
         Else
            If idxTo > UBound(arrayToSort) Then
               Err.Raise ERR_NUM_INVALID_BOUNDARY, _
                         ERR_STR_CLASS_NAME, _
                         ERR_STR_INVALID_BOUNDARY_LEFT & "right" & ERR_STR_INVALID_BOUNDARY_RIGHT
            Else
               InsertionSortWithBoundaries arrayToSort, idxFrom, idxTo
            End If
         End If
      End If
   Else
      Err.Raise ERR_NUM_NO_ARRAY, _
          ERR_STR_CLASS_NAME, _
          ERR_STR_NO_ARRAY
   End If
End Sub

'
'+--------------------------------------------------------------------------
'| Method           | Sort
'|------------------+-------------------------------------------------------
'| Description      | Sort an array of any data type with insertion sort
'|------------------+-------------------------------------------------------
'| Parameter        | arrayToSort: Array to sort
'|------------------+-------------------------------------------------------
'| Return values    | ./.
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2018-09-25  Created. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | The array that is supplied as the parameter is
'|                  | changed so that its elements are sorted.
'+--------------------------------------------------------------------------
'
Public Sub Sort(ByRef arrayToSort As Variant)
   SortPart arrayToSort, LBound(arrayToSort), UBound(arrayToSort)
End Sub
