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
'| Version             | 2.0.0
'|---------------------+---------------------------------------------------
'| Changes             | 2020-09-25  Created. fhs
'|                     | 2020-09-26  Parameters are checked. fhs
'|                     | 2020-09-27  Completely reworked interface. fhs
'|---------------------+---------------------------------------------------
'| Remarks             | The array that is supplied as the parameter is
'|                     | not sorted. Instead an array of indices into the
'|                     | array to sort is created and the array indices are
'|                     | sorted according to the array to sort.
'|                     |
'|                     | If one wants the array elements in sorted order one
'|                     | has to access them like this:
'|                     |
'|                     | For i = LBound(dataArray) To UBound(dataArray)
'|                     |    Debug.Print dataArray(indexArray(i))
'|                     | Next i
'|                     |
'|                     | This method is especially useful if the data to be
'|                     | sorted is large and the cost of moving the data in
'|                     | memory is high, like e.g. for strings.
'|                     |
'|                     | With this method the data is not moved at all and
'|                     | one only needs to access the elements through the
'|                     | index array to get them in sorted order.
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

Private Const ERR_NUM_INVALID_BOUNDARY As Long = ERR_NUM_START + 1
Private Const ERR_STR_INVALID_BOUNDARY_LEFT As String = "Invalid "
Private Const ERR_STR_INVALID_BOUNDARY_RIGHT As String = " boundary"

Private Const ERR_NUM_INVALID_INDEX_ARRAY_BOUNDARY As Long = ERR_NUM_START + 2
Private Const ERR_STR_INVALID_INDEX_ARRAY_BOUNDARY As String = "Boundaries of index array do not match those of the array to sort"

'
' Private methods
'

'
'+--------------------------------------------------------------------------
'| Method           | CheckIsArray
'|------------------+-------------------------------------------------------
'| Description      | Check if supplied paramater is an array
'|------------------+-------------------------------------------------------
'| Parameter        | anArray: Variable to check
'|------------------+-------------------------------------------------------
'| Return values    | ./.
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2020-09-27  Created. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | This method raises an error if the supplied variable
'|                  | is not an array.
'+--------------------------------------------------------------------------
'
Private Sub CheckIsArray(ByRef anArray As Variant)
   If Not IsArray(anArray) Then _
      Err.Raise ERR_NUM_NO_ARRAY, _
                ERR_STR_CLASS_NAME, _
                ERR_STR_NO_ARRAY
End Sub

'
'+--------------------------------------------------------------------------
'| Method           | CheckIndicesAgainstArrayBoundaries
'|------------------+-------------------------------------------------------
'| Description      | Check if supplied paramater is an array
'|------------------+-------------------------------------------------------
'| Parameter        | anArray: Array to check
'|                  | idxFrom: Start index into the array
'|                  | idxTo  : End   index into the array
'|------------------+-------------------------------------------------------
'| Return values    | ./.
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2020-09-27  Created. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | This method raises an error if idxFrom or idxTo are
'|                  | outside the array boundaries.
'+--------------------------------------------------------------------------
'
Private Sub CheckIndicesAgainstArrayBoundaries(ByRef anArray As Variant, ByVal idxFrom As Long, ByVal idxTo As Long)
   If idxFrom < LBound(anArray) Then _
      Err.Raise ERR_NUM_INVALID_BOUNDARY, _
                ERR_STR_CLASS_NAME, _
                ERR_STR_INVALID_BOUNDARY_LEFT & "left" & ERR_STR_INVALID_BOUNDARY_RIGHT

   If idxTo > UBound(anArray) Then _
      Err.Raise ERR_NUM_INVALID_BOUNDARY, _
                ERR_STR_CLASS_NAME, _
                ERR_STR_INVALID_BOUNDARY_LEFT & "right" & ERR_STR_INVALID_BOUNDARY_RIGHT
End Sub

'
'+--------------------------------------------------------------------------
'| Method           | CheckArrayBoundariesMatch
'|------------------+-------------------------------------------------------
'| Description      | Check if the boundaries of the supplied arrays match
'|------------------+-------------------------------------------------------
'| Parameter        | anArray: Array to check
'|                  | idxFrom: Start index into the array
'|                  | idxTo  : End   index into the array
'|------------------+-------------------------------------------------------
'| Return values    | ./.
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2020-09-27  Created. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | This method raises an error if the boundaries of the
'|                  | supplied arrays do not match.
'+--------------------------------------------------------------------------
'
Private Sub CheckArrayBoundariesMatch(ByRef anArray As Variant, ByRef indexArray() As Long)
   If (LBound(indexArray) <> LBound(anArray)) Or _
      (UBound(indexArray) <> UBound(anArray)) Then _
      Err.Raise ERR_NUM_INVALID_INDEX_ARRAY_BOUNDARY, _
                ERR_STR_CLASS_NAME, _
                ERR_STR_INVALID_INDEX_ARRAY_BOUNDARY
End Sub

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
'| Method           | CreateIndexArray
'|------------------+-------------------------------------------------------
'| Description      | Initialize the index array to be used for sorting
'|------------------+-------------------------------------------------------
'| Parameter        | arrayToSort: Array to sort
'|------------------+-------------------------------------------------------
'| Return values    | Index array with indices into the array to sort
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2018-09-25  Created. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | ./.
'+--------------------------------------------------------------------------
'
Private Function CreateIndexArray(ByRef arrayToSort As Variant) As Long()
   Dim result() As Long
   
   Dim i As Long
   Dim j As Long
   Dim low As Long
   Dim Size As Long

   low = LBound(arrayToSort) - 1
   Size = UBound(arrayToSort) - low

   ReDim result(1 To Size)

   j = low
   For i = 1 To Size
      j = j + 1
      result(i) = j
   Next i

   CreateIndexArray = result
End Function

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
'| Method           | SortPartWithExistingIndex
'|------------------+-------------------------------------------------------
'| Description      | Sort part of an array of any data type with an
'|                  | existing index array with optimized QuickSort
'|------------------+-------------------------------------------------------
'| Parameter        | arrayToSort: Array to sort
'|                  | indexArray : Existing array of indicices into the
'|                  |              array to sort
'|                  | idxFrom    : Start index into the array
'|                  | idxTo      : End   index into the array
'|------------------+-------------------------------------------------------
'| Return values    | ./.
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2018-09-27  Created. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | ./.
'+--------------------------------------------------------------------------
'
Public Sub SortPartWithExistingIndex(ByRef arrayToSort As Variant, ByRef indexArray() As Long, ByVal idxFrom As Long, ByVal idxTo As Long)
   CheckIsArray arrayToSort

   Dim arrayLength As Long
   arrayLength = GetArrayLength(arrayToSort)

   If arrayLength > 1 Then
      CheckArrayBoundariesMatch arrayToSort, indexArray
      CheckIndicesAgainstArrayBoundaries arrayToSort, idxFrom, idxTo

      InsertionSortWithIndexArrayAndBoundaries arrayToSort, indexArray, idxFrom, idxTo
   End If
End Sub

'
'+--------------------------------------------------------------------------
'| Method           | SortWithExistingIndex
'|------------------+-------------------------------------------------------
'| Description      | Sort an array of any data type with an exisitng index
'|                  | array with optimized QuickSort
'|------------------+-------------------------------------------------------
'| Parameter        | arrayToSort: Array to sort
'|                  | indexArray : Existing array of indicices into the
'|                  |              array to sort
'|------------------+-------------------------------------------------------
'| Return values    | ./.
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2018-09-27  Created. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | ./.
'+--------------------------------------------------------------------------
'
Public Sub SortWithExistingIndex(ByRef arrayToSort As Variant, ByRef indexArray() As Long)
   CheckIsArray arrayToSort

   Dim arrayLength As Long
   arrayLength = GetArrayLength(arrayToSort)

   If arrayLength > 1 Then
      CheckArrayBoundariesMatch arrayToSort, indexArray

      InsertionSortWithIndexArrayAndBoundaries arrayToSort, indexArray, LBound(arrayToSort), UBound(arrayToSort)
   End If
End Sub

'
'+--------------------------------------------------------------------------
'| Method           | SortPartWithNewIndex
'|------------------+-------------------------------------------------------
'| Description      | Sort part of an array of any data type with an
'|                  | index array with optimized QuickSort and return
'|                  | the created index array
'|------------------+-------------------------------------------------------
'| Parameter        | arrayToSort: Array to sort
'|                  | idxFrom    : Start index into the array
'|                  | idxTo      : End   index into the array
'|------------------+-------------------------------------------------------
'| Return values    | Index array with indices into arrayToSort
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2018-09-25  Created. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | ./.
'+--------------------------------------------------------------------------
'
Public Function SortPartWithNewIndex(ByRef arrayToSort As Variant, ByVal idxFrom As Long, ByVal idxTo As Long) As Long()
   CheckIsArray arrayToSort

   Dim arrayLength As Long
   arrayLength = GetArrayLength(arrayToSort)

   Dim indexArray() As Long

   If arrayLength > 0 Then
      CheckIndicesAgainstArrayBoundaries arrayToSort, idxFrom, idxTo

      indexArray = CreateIndexArray(arrayToSort)

      InsertionSortWithIndexArrayAndBoundaries arrayToSort, indexArray, idxFrom, idxTo
   End If

   SortPartWithNewIndex = indexArray
End Function

'
'+--------------------------------------------------------------------------
'| Method           | SortWithNewIndex
'|------------------+-------------------------------------------------------
'| Description      | Sort an array of any data type with an
'|                  | index array with optimized QuickSort and return
'|                  | the created index array
'|------------------+-------------------------------------------------------
'| Parameter        | arrayToSort: Array to sort
'|------------------+-------------------------------------------------------
'| Return values    | Index array with indices into arrayToSort
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2018-09-27  Created. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | ./.
'+--------------------------------------------------------------------------
'
Public Function SortWithNewIndex(ByRef arrayToSort As Variant) As Long()
   CheckIsArray arrayToSort

   Dim arrayLength As Long
   arrayLength = GetArrayLength(arrayToSort)

   Dim indexArray() As Long

   If arrayLength > 0 Then
      indexArray = CreateIndexArray(arrayToSort)

      InsertionSortWithIndexArrayAndBoundaries arrayToSort, indexArray, LBound(arrayToSort), UBound(arrayToSort)
   End If

   SortWithNewIndex = indexArray
End Function
