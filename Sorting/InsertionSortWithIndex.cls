VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "InsertionSortWithIndex"
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
'| Class               | InsertionSortWithIndex
'|---------------------+---------------------------------------------------
'| Description         | Insertion sort implementation with an index array
'|---------------------+---------------------------------------------------
'| Author              | Frank Schwab
'|---------------------+---------------------------------------------------
'| Version             | 1.0.0
'|---------------------+---------------------------------------------------
'| Changes             | 2020-07-20  Created. fhs
'|---------------------+---------------------------------------------------
'

Option Explicit

'
' Constants for errors
'
Private Const ERR_STR_CLASS_NAME As String = "InsertionSortWithIndex"

Private Const ERR_NUM_START As Long = vbObjectError + 55328

Private Const ERR_NUM_NO_ARRAY As Long = ERR_NUM_START
Private Const ERR_STR_NO_ARRAY As String = "Supplied parameter is not an array"

'
' Private methods
'

'
'+--------------------------------------------------------------------------
'| Method           | InitializeIndexArray
'|------------------+-------------------------------------------------------
'| Description      | Initialize the index array to be used for sorting
'|------------------+-------------------------------------------------------
'| Parameter        | arrayToSort: Array to sort
'|                  | indexArray : Index array with indices into the
'|                  | array to sort (this array is created in this method)
'|------------------+-------------------------------------------------------
'| Return values    | ./.
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2018-09-25  Created. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | ./.
'+--------------------------------------------------------------------------
'
Private Sub InitializeIndexArray(ByRef arrayToSort As Variant, ByRef indexArray() As Long)
   Dim i As Long
   Dim j As Long
   Dim low As Long
   Dim Size As Long

   low = LBound(arrayToSort) - 1
   Size = UBound(arrayToSort) - low

   ReDim indexArray(1 To Size)

   j = low
   For i = 1 To Size
      j = j + 1
      indexArray(i) = j
   Next i
End Sub

'
'+--------------------------------------------------------------------------
'| Method           | InsertionSortWithIndexArrayAndBoundaries
'|------------------+-------------------------------------------------------
'| Description      | Sort an array of any data type with an index array
'|                  | with insertion sort
'|------------------+-------------------------------------------------------
'| Parameter        | arrayToSort: Array to sort
'|                  | indexArray : Index array with indices into the
'|                  | array to sort (this array is created in this method)
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
'|                  | not sorted. Instead an array of indices into the
'|                  | array to sort is created and the array indices are
'|                  | sorted according to the array to sort.
'|                  |
'|                  | If one wants the array elements in sorted order one
'|                  | has to access them like this:
'|                  |
'|                  | For i = LBound(dataArray) To UBound(dataArray)
'|                  |    Debug.Print dataArray(indexArray(i))
'|                  | Next i
'|                  |
'|                  | This method is especially useful if the data to be
'|                  | sorted is large and the cost of moving the data in
'|                  | memory is high, like e.g. for strings.
'|                  |
'|                  | With this method the data is not moved at all and
'|                  | one only needs to access the elements through the
'|                  | index array to get them in sorted order.
'+--------------------------------------------------------------------------
'
Private Sub InsertionSortWithIndexArrayAndBoundaries(ByRef arrayToSort As Variant, _
                                                     ByRef indexArray() As Long, _
                                                     ByVal idxFrom As Long, _
                                                     ByVal idxTo As Long)
   Dim idxLeft As Long
   Dim idxLeftForTest As Long
   Dim idxRight As Long
   Dim valueToInsert As Variant
   Dim indexToInsert As Long
   Dim compareIndex As Long

   For idxRight = idxFrom + 1 To idxTo
      indexToInsert = indexArray(idxRight)
      valueToInsert = arrayToSort(indexToInsert)
      idxLeft = idxRight

      Do
         idxLeftForTest = idxLeft - 1

         If idxLeftForTest >= idxFrom Then
            compareIndex = indexArray(idxLeftForTest)

            If arrayToSort(compareIndex) > valueToInsert Then
               indexArray(idxLeft) = compareIndex
               idxLeft = idxLeftForTest
            Else
               Exit Do
            End If
         Else
            Exit Do
         End If
      Loop

      indexArray(idxLeft) = indexToInsert
   Next idxRight
End Sub

'
' Public methods
'

'
'+--------------------------------------------------------------------------
'| Method           | SortPartWithIndex
'|------------------+-------------------------------------------------------
'| Description      | Sort a part of an array of any data type with an
'|                  | index array with insertion sort
'|------------------+-------------------------------------------------------
'| Parameter        | arrayToSort: Array to sort
'|                  | indexArray : Index array with indices into the
'|                  |              array to sort (this array is created
'|                  |              in this method)
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
Public Sub SortPartWithIndex(ByRef arrayToSort As Variant, ByRef indexArray() As Long, ByVal idxFrom As Long, ByVal idxTo As Long)
   If IsArray(arrayToSort) Then
      InitializeIndexArray arrayToSort, indexArray
      InsertionSortWithIndexArrayAndBoundaries arrayToSort, indexArray, idxFrom, idxTo
   Else
      Err.Raise ERR_NUM_NO_ARRAY, _
          ERR_STR_CLASS_NAME, _
          ERR_STR_NO_ARRAY
   End If
End Sub

'
'+--------------------------------------------------------------------------
'| Method           | SortWithIndex
'|------------------+-------------------------------------------------------
'| Description      | Sort an array of any data type with an index array
'|                  | with insertion sort
'|------------------+-------------------------------------------------------
'| Parameter        | arrayToSort: Array to sort
'|                  | indexArray : Index array with indices into the
'|                  |              array to sort (this array is created
'|                  |              in this method)
'|------------------+-------------------------------------------------------
'| Return values    | ./.
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2018-09-25  Created. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | The array that is supplied as the parameter is
'|                  | not sorted. Instead an array of indices into the
'|                  | array to sort is created and the array indices are
'|                  | sorted according to the array to sort.
'|                  |
'|                  | If one wants the array elements in sorted order one
'|                  | has to access them like this:
'|                  |
'|                  | For i = LBound(dataArray) To UBound(dataArray)
'|                  |    Debug.Print dataArray(indexArray(i))
'|                  | Next i
'|                  |
'|                  | This method is especially useful if the data to be
'|                  | sorted is large and the cost of moving the data in
'|                  | memory is high, like e.g. for strings.
'|                  |
'|                  | With this method the data is not moved at all and
'|                  | one only needs to access the elements through the
'|                  | index array to get them in sorted order.
'+--------------------------------------------------------------------------
'
Public Sub SortWithIndex(ByRef arrayToSort As Variant, ByRef indexArray() As Long)
   SortPartWithIndex arrayToSort, indexArray, LBound(arrayToSort), UBound(arrayToSort)
End Sub
