VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "QuickSort"
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
'| Class               | QuickSort
'|---------------------+---------------------------------------------------
'| Description         | Optimized quicksort implementation
'|---------------------+---------------------------------------------------
'| Author              | Frank Schwab
'|---------------------+---------------------------------------------------
'| Version             | 1.0.0
'|---------------------+---------------------------------------------------
'| Changes             | 2020-09-25  Created. fhs
'|---------------------+---------------------------------------------------
'| Remarks             | This is an optimized QuickSort implementation
'|                     | that is combined with InsertionSort.
'|---------------------+---------------------------------------------------
'

Option Explicit

'
' Constants for errors
'
Private Const ERR_STR_CLASS_NAME As String = "QuickSort"

Private Const ERR_NUM_START As Long = vbObjectError + 50665

Private Const ERR_NUM_NO_ARRAY As Long = ERR_NUM_START
Private Const ERR_STR_NO_ARRAY As String = "Supplied parameter is not an array"

'
' Private constants
'
Private Const CUTOFF_SIZE_MINUS_1 As Long = 9   ' Cut off if 10 elements or less to sort, so set to 10 - 1 = 9

'
' Private methods
'

'
'+--------------------------------------------------------------------------
'| Method           | QuicksortWithBoundaries
'|------------------+-------------------------------------------------------
'| Description      | Sort an array of any data type with optimized QuickSort
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
'|                  | changed so that its elements are sorted.
'+--------------------------------------------------------------------------
'
Private Sub QuicksortWithBoundaries(ByRef arrayToSort As Variant, ByVal idxFrom As Long, ByVal idxTo As Long)
   Dim pivot As Variant
   Dim aStack As New Stack
   Dim idxCenter As Long
   Dim idxPartitionTo As Long
   Dim idxLeft As Long
   Dim idxRight As Long
   Dim parkElement As Variant
   Dim leftSize As Long
   Dim rightSize As Long
   Dim insertionSorter As New InsertionSort

'
' The whole thing is pushed to the stack as a dummy. If it is popped from the stack
' the sorting just finishes. It is not executed again with these boundaries.
' One could just push dummy values, as well.
'
   aStack.Push idxFrom
   aStack.Push idxTo

   Do
' Only do quicksort if the partition is larger then the cutoff size
      If (idxTo - idxFrom) >= CUTOFF_SIZE_MINUS_1 Then
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

         If arrayToSort(idxFrom) > arrayToSort(idxCenter) Then
            parkElement = arrayToSort(idxFrom)
            arrayToSort(idxFrom) = arrayToSort(idxCenter)
            arrayToSort(idxCenter) = parkElement
         End If

         If arrayToSort(idxFrom) > arrayToSort(idxTo) Then
            parkElement = arrayToSort(idxFrom)
            arrayToSort(idxFrom) = arrayToSort(idxTo)
            arrayToSort(idxTo) = parkElement
         End If

         If arrayToSort(idxCenter) > arrayToSort(idxTo) Then
            parkElement = arrayToSort(idxCenter)
            arrayToSort(idxCenter) = arrayToSort(idxTo)
            arrayToSort(idxTo) = parkElement
         End If

         pivot = arrayToSort(idxCenter)

         idxPartitionTo = idxTo - 1
         If idxCenter <> idxPartitionTo Then
            arrayToSort(idxCenter) = arrayToSort(idxPartitionTo)
            arrayToSort(idxPartitionTo) = pivot
         End If

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
' Here we can safely assume that idxRight is greater than idxLeft since the
' CUTOFF_SIZE_MINUS_1 will always be larger than 3.
'
         Do
            Do
               idxLeft = idxLeft + 1
            Loop Until arrayToSort(idxLeft) >= pivot
      
            Do
               idxRight = idxRight - 1
            Loop Until arrayToSort(idxRight) <= pivot

'
' Exchange the elements at indices idxLeft and idxRight,
' but only if the indices haven't crossed yet.
' If they crossed, just end the loop and don't exchange.
'
            parkElement = arrayToSort(idxLeft)
            If idxLeft < idxRight Then
               arrayToSort(idxLeft) = arrayToSort(idxRight)
               arrayToSort(idxRight) = parkElement
            Else
               Exit Do
            End If
         Loop

'
' Now put the pivot in the correct place at idxLeft
'
         arrayToSort(idxLeft) = arrayToSort(idxPartitionTo)
         arrayToSort(idxPartitionTo) = parkElement

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
' Use insertion sort if the partition is smaller than or equal to the cutoff size
         If idxTo > idxFrom Then _
            insertionSorter.SortPart arrayToSort, idxFrom, idxTo

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
'| Method           | SortPart
'|------------------+-------------------------------------------------------
'| Description      | Sort a part of an array of any data type
'|                  | with optimized QuickSort
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
      QuicksortWithBoundaries arrayToSort, idxFrom, idxTo
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
'| Description      | Sort an array of any data type with optimized QuickSort
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
