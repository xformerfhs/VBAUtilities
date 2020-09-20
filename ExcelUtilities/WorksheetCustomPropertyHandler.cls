VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WorksheetCustomPropertyHandler"
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
'| Class               | WorksheetCustomPropertyHandler
'|---------------------+---------------------------------------------------
'| Description         | Wrapper class for managing Excel worksheet
'|                     | custom properties
'|---------------------+---------------------------------------------------
'| Author              | Frank Schwab
'|---------------------+---------------------------------------------------
'| Version             | 1.0.0
'|---------------------+---------------------------------------------------
'| Changes             | 2018-09-25  Created. fhs
'|---------------------+---------------------------------------------------
'| Remarks             | The MS supplied API for custom properties is 
'|                     | cumbersome and incomplete. This wrapper class
'|                     | makes custom property handling much easier.
'|---------------------+---------------------------------------------------
'| Typical usage       | Dim wscph As New WorksheetCustomPropertyHandler
'|                     | ...
'|                     | Dim lastModDate As Variant
'|                     | lastModDate = wscph.SafeGetCustomProperty(ActiveWorksheet, "LastModificationDate")
'|                     | ...
'|                     | wscph.SetCustomProperty(ActiveWorksheet, "LastModificationDate", newDate)
'+-------------------------------------------------------------------------

Option Explicit

'
' Error constants
'
Private Const ERR_STR_CLASS_NAME As String = "WorksheetCustomPropertyHandler"
Private Const ERR_NUM_START As Long = 27792

Private Const ERR_NUM_NOT_FOUND As Long = ERR_NUM_START
Private Const ERR_STR_NOT_FOUND As String = "There is no property named '"


'
' Helper constant for StrComp
'
Private Const STRCOMP_EQUAL As Integer = 0


'
' Private methods
'

'
'+--------------------------------------------------------------------------
'| Method           | IsEmptyOrNull
'|------------------+-------------------------------------------------------
'| Description      | Check wether the variant data type is empty or null
'|------------------+-------------------------------------------------------
'| Parameter        | propertyValue: Variant variable to check
'|------------------+-------------------------------------------------------
'| Return values    | True : Variable is empty or null
'|                  | False: Variable is not empty or null
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2018-09-25  Created. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | ./.
'+--------------------------------------------------------------------------
'
Private Function IsEmptyOrNull(ByRef propertyValue As Variant) As Boolean
   IsEmptyOrNull = (VarType(propertyValue) <= 1)
End Function
   

'
' Public methods
'

'
'+--------------------------------------------------------------------------
'| Method           | CustomPropertyExists
'|------------------+-------------------------------------------------------
'| Description      | Checks wether a property with a given name exists
'|------------------+-------------------------------------------------------
'| Parameter        | propertyWorksheet: The worksheet for the property
'|                  | propertyName     : The property name
'|------------------+-------------------------------------------------------
'| Return values    | True : A property with the specified name exists
'|                  | False: A property with the specified name does not exist
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2018-09-25  Created. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | Property name comparisons are case-insensitive!
'|                  | A property with the name "aProperty" is the same as
'|                  | a property with the name "ApRoPeRtY". This is the way
'|                  | Excel works.
'+--------------------------------------------------------------------------
'
Public Function CustomPropertyExists(ByVal propertyWorksheet As Worksheet, ByRef propertyName As String) As Boolean
   Dim aProperty As CustomProperty

   CustomPropertyExists = False

   For Each aProperty In propertyWorksheet.CustomProperties
      If StrComp(aProperty.Name, propertyName, vbTextCompare) = STRCOMP_EQUAL Then
         CustomPropertyExists = True

         Exit Function
      End If
   Next
End Function

'
'+--------------------------------------------------------------------------
'| Method           | GetCustomProperty
'|------------------+-------------------------------------------------------
'| Description      | Get the value of a custom property
'|------------------+-------------------------------------------------------
'| Parameter        | propertyWorksheet: The worksheet for the property
'|                  | propertyName     : The property name
'|------------------+-------------------------------------------------------
'| Return values    | Value of the custom property
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2018-09-25  Created. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | This method raises an error if a custom property
'|                  | with the specified name does not exist.
'+--------------------------------------------------------------------------
'
Public Function GetCustomProperty(ByRef propertyWorksheet As Worksheet, ByRef propertyName As String) As Variant
   If CustomPropertyExists(propertyWorksheet, propertyName) Then
      GetCustomProperty = propertyWorksheet.CustomProperties.Item(propertyName).value
   Else
      Err.Raise ERR_NUM_NOT_FOUND, _
                ERR_STR_CLASS_NAME, _
                ERR_STR_NOT_FOUND & _
                   propertyName & _
                   "'"
   End If
End Function

'
'+--------------------------------------------------------------------------
'| Method           | SafeGetCustomProperty
'|------------------+-------------------------------------------------------
'| Description      | Get the value of a custom property
'|------------------+-------------------------------------------------------
'| Parameter        | propertyWorksheet: The worksheet for the property
'|                  | propertyName     : The property name
'|------------------+-------------------------------------------------------
'| Return values    | Value of the custom property
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2018-09-25  Created. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | This method does not raise an error if a custom property
'|                  | with the specified name does not exist. An empty
'|                  | value is returned in this case.
'+--------------------------------------------------------------------------
'
Public Function SafeGetCustomProperty(ByRef propertyWorksheet As Worksheet, ByRef propertyName As String) As Variant
   If CustomPropertyExists(propertyWorksheet, propertyName) Then
      SafeGetCustomProperty = propertyWorksheet.CustomProperties.Item(propertyName).value
   Else
      SafeGetCustomProperty = Empty
   End If
End Function

'
'+--------------------------------------------------------------------------
'| Method           | SetCustomProperty
'|------------------+-------------------------------------------------------
'| Description      | Set a custom property to a value
'|------------------+-------------------------------------------------------
'| Parameter        | propertyWorksheet: The worksheet for the property
'|                  | propertyName     : The property name
'|                  | propertyValue    : The property value
'|------------------+-------------------------------------------------------
'| Return values    | ./.
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2018-09-25  Created. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | This method shows how incredibly complicated
'|                  | the handling of custom properties in Excel is.
'+--------------------------------------------------------------------------
'
Public Sub SetCustomProperty(ByRef propertyWorksheet As Worksheet, ByRef propertyName As String, ByVal propertyValue As Variant)
   Dim actualProperty As CustomProperty

   '
   ' Setting a property value is quite complicated. One has to check wether
   ' the property already exists, or not. One has to check for null
   ' or empty.
   '

   '
   ' Check, if property already exists?
   '
   If CustomPropertyExists(propertyWorksheet, propertyName) Then
      '
      ' The property already exists
      '
      ' So get it
      '
      Set actualProperty = propertyWorksheet.CustomProperties.Item(propertyName)

      If IsEmptyOrNull(propertyValue) Then
         '
         ' Delete it if the new value is empty or null
         '
         actualProperty.Delete
      Else
         '
         ' The new value is neither empty, nor null. So set it.
         '

         actualProperty.value = propertyValue
      End If
   Else
      '
      ' The property does not exist
      '
      ' Just add it (That was the easy part)
      '
      propertyWorksheet.CustomProperties.Add Name:=propertyName, value:=propertyValue
   End If
End Sub

'
'+--------------------------------------------------------------------------
'| Method           | DeleteProperty
'|------------------+-------------------------------------------------------
'| Description      | Delete a custom property
'|------------------+-------------------------------------------------------
'| Parameter        | propertyWorksheet: The worksheet for the property
'|                  | propertyName     : The property name
'|------------------+-------------------------------------------------------
'| Return values    | ./.
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2018-09-25  Created. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | This method raises an error if a custom property
'|                  | with the specified name does not exist.
'+--------------------------------------------------------------------------
'
Public Sub DeleteProperty(ByRef propertyWorksheet As Worksheet, ByRef propertyName As String)
   If CustomPropertyExists(propertyWorksheet, propertyName) Then
      propertyWorksheet.CustomProperties.Item(propertyName).Delete
   Else
      Err.Raise ERR_NUM_NOT_FOUND, _
                ERR_STR_CLASS_NAME, _
                ERR_STR_NOT_FOUND & _
                   propertyName & _
                   "'"
   End If
End Sub

'
'+--------------------------------------------------------------------------
'| Method           | SafeDeleteProperty
'|------------------+-------------------------------------------------------
'| Description      | Delete a custom property
'|------------------+-------------------------------------------------------
'| Parameter        | propertyWorksheet: The worksheet for the property
'|                  | propertyName     : The property name
'|------------------+-------------------------------------------------------
'| Return values    | ./.
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2018-09-25  Created. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | This method does not raise an error if a custom property
'|                  | with the specified name does not exist.
'+--------------------------------------------------------------------------
'
Public Sub SafeDeleteProperty(ByRef propertyWorksheet As Worksheet, ByRef propertyName As String)
   If CustomPropertyExists(propertyWorksheet, propertyName) Then
      propertyWorksheet.CustomProperties.Item(propertyName).Delete
   End If
End Sub
