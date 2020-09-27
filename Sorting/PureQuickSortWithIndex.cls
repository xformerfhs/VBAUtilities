VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PureQuickSortWithIndex"
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
'| Class               | PureQuickSortWithIndex
'|---------------------+---------------------------------------------------
'| Description         | Pure quicksort implementation with an index array
'|---------------------+---------------------------------------------------
'| Author              | Frank Schwab
'|---------------------+---------------------------------------------------
'| Version             | 2.0.0
'|---------------------+---------------------------------------------------
'| Changes             | 2020-09-25  Created. fhs
'|                     | 2020-09-26  Parameters are checked. fhs
'|                     | 2020-09-27  Completely reworked interface. fhs
'|---------------------+---------------------------------------------------
'| Remarks             | This is a pure QuickSort implementation.
'|                     |
'|                     | The array that is supplied as the parameter is
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
Private Const ERR_STR_CLASS_NAME As String = "PureQuickSortWithIndex"

Private Const ERR_NUM_START As Long = vbObjectError + 57421

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
'| Method           | PureQuickSortWithIndexArrayAndBoundaries
'|------------------+-------------------------------------------------------
'| Description      | Sort an array of any data type with an index array
'|                  | with pure QuickSort
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
Private Sub PureQuickSortWithIndexArrayAndBoundaries(ByRef arrayToSort As Variant, _
                                                     ByRef indexArray() As Long, _
                                                     ByVal idxFrom As Long, _
                                                     ByVal idxTo As Long)
   Dim pivot As Variant
   Dim aStack As New Stack
   Dim idxCenter As Long
   Dim idxPartitionTo As Long
   Dim idxLeft As Long
   Dim idxRight As Long
   Dim parkElement As Long
   Dim leftSize As Long
   Dim rightSize As Long

'
' The whole thing is pushed to the stack as a dummy. If it is popped from the stack
' the sorting just finishes. It is not executed again with these boundaries.
' One could just push dummy values, as well.
'
   aStack.Push idxFrom
   aStack.Push idxTo

   Do
'      Debug.Print "("; Format$(aStack.depth); ": "; Format$(idxFrom); ", "; Format$(idxTo); ")"
      If idxTo > idxFrom Then
'
' Median of 3
'
' Choose the median of the first, middle and last array element as the pivot.
' As a side effect these three elements are already sorted.
'
' The pivot is placed at the next-to-last position so that the Quicksort loop
' only has to partition from idxFrom + 1 to idxTo - 1
'
' Normally this would be implemented as a function but has been inlined for better performance.
' As a side effect the correct index for the rightmost array element to partition is already
' computed.
'
         idxCenter = idxFrom + ((idxTo - idxFrom) \ 2)

         If arrayToSort(indexArray(idxFrom)) > arrayToSort(indexArray(idxCenter)) Then
            parkElement = indexArray(idxFrom)
            indexArray(idxFrom) = indexArray(idxCenter)
            indexArray(idxCenter) = parkElement
         End If

         If arrayToSort(indexArray(idxFrom)) > arrayToSort(indexArray(idxTo)) Then
            parkElement = indexArray(idxFrom)
            indexArray(idxFrom) = indexArray(idxTo)
            indexArray(idxTo) = parkElement
         End If

         If arrayToSort(indexArray(idxCenter)) > arrayToSort(indexArray(idxTo)) Then
            parkElement = indexArray(idxCenter)
            indexArray(idxCenter) = indexArray(idxTo)
            indexArray(idxTo) = parkElement
         End If

         parkElement = indexArray(idxCenter)

         idxPartitionTo = idxTo - 1
         If idxCenter <> idxPartitionTo Then
            indexArray(idxCenter) = indexArray(idxPartitionTo)
            indexArray(idxPartitionTo) = parkElement
         End If

         pivot = arrayToSort(indexArray(idxPartitionTo))

'
' Partition
'
' Now the partition loop is run if there are more than 3 elements left
'
' idxLeft is set to the index of the first element to sort which is already sorted.
' The partition loop increments the index before each test, so the first tested element
' is the one with the index idxFrom + 1.
'
' In the same sense, idxRight is set to the index of the pivot element. The partition loop
' decrements the index before each test, so the first tested element is the one
' with the index idxPartitionTo - 1.
'
         idxLeft = idxFrom
         idxRight = idxPartitionTo

'
' Only start Quicksort if we have more than three elements, so the difference between idxLeft
' and idxRight has to be *greater* than 1. A difference of 1 indicates a partition size of 3.
'
         If (idxRight - idxLeft) > 1 Then
            Do
               Do
                  idxLeft = idxLeft + 1
               Loop Until arrayToSort(indexArray(idxLeft)) >= pivot

               Do
                  idxRight = idxRight - 1
               Loop Until arrayToSort(indexArray(idxRight)) <= pivot

'
' Exchange the elements at indices idxLeft and idxRight,
' but only if the indices haven't crossed yet.
' If they crossed, just end the loop and don't exchange.
'
               parkElement = indexArray(idxLeft)
               If idxLeft < idxRight Then
                  indexArray(idxLeft) = indexArray(idxRight)
                  indexArray(idxRight) = parkElement
               Else
                  Exit Do
               End If
            Loop

'
' Now put the pivot in the correct place at idxLeft
'
            indexArray(idxLeft) = indexArray(idxPartitionTo)
            indexArray(idxPartitionTo) = parkElement

'
' Now check which part to sort next.
' Calculate the size of the parts to the left and to the right of idxLeft element
' push the larger part on the stack and immediately sort the smaller part
' by reassigning the boundaries.
'
            leftSize = idxLeft - idxFrom
            rightSize = idxTo - idxLeft
            If leftSize > rightSize Then
               aStack.Push idxFrom
               aStack.Push idxLeft - 1

               idxFrom = idxLeft + 1
            Else
               aStack.Push idxLeft + 1
               aStack.Push idxTo

               idxTo = idxLeft - 1
            End If
         Else
            idxTo = aStack.Pop
            idxFrom = aStack.Pop
         End If
      Else
         idxTo = aStack.Pop
         idxFrom = aStack.Pop
      End If
   Loop Until aStack.IsEmpty
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

      PureQuickSortWithIndexArrayAndBoundaries arrayToSort, indexArray, idxFrom, idxTo
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

      PureQuickSortWithIndexArrayAndBoundaries arrayToSort, indexArray, LBound(arrayToSort), UBound(arrayToSort)
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

      PureQuickSortWithIndexArrayAndBoundaries arrayToSort, indexArray, idxFrom, idxTo
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

      PureQuickSortWithIndexArrayAndBoundaries arrayToSort, indexArray, LBound(arrayToSort), UBound(arrayToSort)
   End If

   SortWithNewIndex = indexArray
End Function
