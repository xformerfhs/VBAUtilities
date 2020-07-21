VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sorter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
'+-------------------------------------------------------------------------
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
'| Class               | Sorter
'|---------------------+---------------------------------------------------
'| Description         | Several optimized sorting algorithms
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
' Private constants
'
Private Const CUTOFF_SIZE_MINUS_1 As Long = 9   ' Cut off if 10 elements or less to sort, so set to 10 - 1 = 9

'
' Instance variables
'

'
' Private methods
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

Private Sub PureQuicksortPrivate(ByRef arrayToSort As Variant)
   Dim pivot As Variant
   Dim aStack As New Stack
   Dim idxFrom As Long
   Dim idxTo As Long
   Dim idxCenter As Long
   Dim idxPartitionTo As Long
   Dim idxLeft As Long
   Dim idxRight As Long
   Dim parkElement As Variant
   Dim leftSize As Long
   Dim rightSize As Long

   idxFrom = LBound(arrayToSort)
   idxTo = UBound(arrayToSort)

'
' The whole thing is pushed to the stack as a dummy. If it is popped from the stack
' the sorting just finishes. It is not executed again with these boundaries.
' One could just push dummy values, as well.
'
   aStack.push idxFrom
   aStack.push idxTo

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
' Only start Quicksort if we have more than three elements, so the difference between idxLeft
' and idxRight has to be *greater* than 1. A difference of 1 indicates a partition size of 3.
'
         If (idxRight - idxLeft) > 1 Then
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
               aStack.push idxFrom
               aStack.push idxLeft - 1
            
               idxFrom = idxLeft + 1
            Else
               aStack.push idxLeft + 1
               aStack.push idxTo
            
               idxTo = idxLeft - 1
            End If
         Else
            idxTo = aStack.pop
            idxFrom = aStack.pop
         End If
      Else
         idxTo = aStack.pop
         idxFrom = aStack.pop
      End If
   Loop Until aStack.isEmpty
End Sub

Private Sub QuicksortPrivate(ByRef arrayToSort As Variant)
   Dim pivot As Variant
   Dim aStack As New Stack
   Dim idxFrom As Long
   Dim idxTo As Long
   Dim idxCenter As Long
   Dim idxPartitionTo As Long
   Dim idxLeft As Long
   Dim idxRight As Long
   Dim parkElement As Variant
   Dim leftSize As Long
   Dim rightSize As Long

   idxFrom = LBound(arrayToSort)
   idxTo = UBound(arrayToSort)

'
' The whole thing is pushed to the stack as a dummy. If it is popped from the stack
' the sorting just finishes. It is not executed again with these boundaries.
' One could just push dummy values, as well.
'
   aStack.push idxFrom
   aStack.push idxTo

   Do
'      Debug.Print "("; Format$(aStack.depth); ": "; Format$(idxFrom); ", "; Format$(idxTo); ")"
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
            aStack.push idxFrom
            aStack.push idxLeft - 1

            idxFrom = idxLeft + 1
         Else
            aStack.push idxLeft + 1
            aStack.push idxTo

            idxTo = idxLeft - 1
         End If
      Else
         If idxTo > idxFrom Then
            InsertionSortWithBoundaries arrayToSort, idxFrom, idxTo
         End If

         idxTo = aStack.pop
         idxFrom = aStack.pop
      End If
   Loop Until aStack.isEmpty
End Sub

Private Sub InsertionsortWithIndexArrayAndBoundaries(ByRef arrayToSort As Variant, _
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

Private Sub PureQuicksortWithIndexArrayPrivate(ByRef arrayToSort As Variant, ByRef indexArray() As Long)
   Dim pivot As Variant
   Dim aStack As New Stack
   Dim idxFrom As Long
   Dim idxTo As Long
   Dim idxCenter As Long
   Dim idxPartitionTo As Long
   Dim idxLeft As Long
   Dim idxRight As Long
   Dim parkElement As Long
   Dim leftSize As Long
   Dim rightSize As Long

   idxFrom = LBound(arrayToSort)
   idxTo = UBound(arrayToSort)
   
'
' The whole thing is pushed to the stack as a dummy. If it is popped from the stack
' the sorting just finishes. It is not executed again with these boundaries.
' One could just push dummy values, as well.
'
   aStack.push idxFrom
   aStack.push idxTo

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
               aStack.push idxFrom
               aStack.push idxLeft - 1

               idxFrom = idxLeft + 1
            Else
               aStack.push idxLeft + 1
               aStack.push idxTo

               idxTo = idxLeft - 1
            End If
         Else
            idxTo = aStack.pop
            idxFrom = aStack.pop
         End If
      Else
         idxTo = aStack.pop
         idxFrom = aStack.pop
      End If
   Loop Until aStack.isEmpty
End Sub

Private Sub QuicksortWithIndexArrayPrivate(ByRef arrayToSort As Variant, ByRef indexArray() As Long)
   Dim pivot As Variant
   Dim aStack As New Stack
   Dim idxFrom As Long
   Dim idxTo As Long
   Dim idxCenter As Long
   Dim idxPartitionTo As Long
   Dim idxLeft As Long
   Dim idxRight As Long
   Dim parkElement As Long
   Dim leftSize As Long
   Dim rightSize As Long

   idxFrom = LBound(arrayToSort)
   idxTo = UBound(arrayToSort)

'
' The whole thing is pushed to the stack as a dummy. If it is popped from the stack
' the sorting just finishes. It is not executed again with these boundaries.
' One could just push dummy values, as well.
'
   aStack.push idxFrom
   aStack.push idxTo

   Do
'      Debug.Print "("; Format$(aStack.depth); ": "; Format$(idxFrom); ", "; Format$(idxTo); ")"
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

         idxPartitionTo = idxTo - 1

'
' Here it is safe to assume that idxPartitionTo is not equal to idxCenter
' since the cutoff value will always be greater than 2
'
         parkElement = indexArray(idxCenter)
         indexArray(idxCenter) = indexArray(idxPartitionTo)
         indexArray(idxPartitionTo) = parkElement

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
' Here we can safely assume that idxRight is greater than idxLeft since the
' CUTOFF_SIZE_MINUS_1 will always be larger than 3.
'
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
            aStack.push idxFrom
            aStack.push idxLeft - 1

            idxFrom = idxLeft + 1
         Else
            aStack.push idxLeft + 1
            aStack.push idxTo

            idxTo = idxLeft - 1
         End If
      Else
         If idxTo > idxFrom Then
            InsertionsortWithIndexArrayAndBoundaries arrayToSort, indexArray, idxFrom, idxTo
         End If

         idxTo = aStack.pop
         idxFrom = aStack.pop
      End If
   Loop Until aStack.isEmpty
End Sub

'
' Public methods
'
Public Sub InsertionSort(ByRef arrayToSort As Variant)
   If IsArray(arrayToSort) Then
      InsertionSortWithBoundaries arrayToSort, LBound(arrayToSort), UBound(arrayToSort)
   End If
End Sub

Public Sub PureQuicksort(ByRef arrayToSort As Variant)
   If IsArray(arrayToSort) Then
      PureQuicksortPrivate arrayToSort
   End If
End Sub

Public Sub Quicksort(ByRef arrayToSort As Variant)
   If IsArray(arrayToSort) Then
      QuicksortPrivate arrayToSort
   End If
End Sub

Public Sub InsertionSortWithIndexArray(ByRef arrayToSort As Variant, ByRef indexArray() As Long)
   If IsArray(arrayToSort) Then
      InsertionsortWithIndexArrayAndBoundaries arrayToSort, indexArray, LBound(arrayToSort), UBound(arrayToSort)
   End If
End Sub

Public Sub PureQuicksortWithIndexArray(ByRef arrayToSort As Variant, ByRef indexArray() As Long)
   If IsArray(arrayToSort) Then
      InitializeIndexArray arrayToSort, indexArray
      PureQuicksortWithIndexArrayPrivate arrayToSort, indexArray
   End If
End Sub

Public Sub QuicksortWithIndexArray(ByRef arrayToSort As Variant, ByRef indexArray() As Long)
   If IsArray(arrayToSort) Then
      InitializeIndexArray arrayToSort, indexArray
      QuicksortWithIndexArrayPrivate arrayToSort, indexArray
   End If
End Sub
