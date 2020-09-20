VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WorkbookCustomPropertyHandler"
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
'| Class               | WorkbookCustomPropertyHandler
'|---------------------+---------------------------------------------------
'| Description         | Wrapper class for managing Excel workbook
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
'| Typical usage       | Dim wbcph As New WorkbookCustomPropertyHandler
'|                     | ...
'|                     | Dim lastModDate As Variant
'|                     | lastModDate = wbcph.SafeGetCustomProperty(ActiveWorkbook, "LastModificationDate")
'|                     | ...
'|                     | wbcph.SetCustomProperty(ActiveWorkbook, "LastModificationDate", newDate)
'+-------------------------------------------------------------------------

Option Explicit

'
' Error constants
'
Private Const ERR_STR_CLASS_NAME As String = "WorkbookCustomPropertyHandler"
Private Const ERR_NUM_START As Long = 38803

Private Const ERR_NUM_PROPERTY_NOT_FOUND As Long = ERR_NUM_START
Private Const ERR_STR_PROPERTY_NOT_FOUND As String = "There is no property named '"

Private Const ERR_NUM_TYPE_NOT_MAPPABLE As Long = ERR_NUM_START + 1
Private Const ERR_STR_TYPE_NOT_MAPPABLE As String = "variant data type can not be mapped to document property data type"

Private Const ERR_NUM_TYPE_MISMATCH As Long = ERR_NUM_START + 2
Private Const ERR_STR_TYPE_MISMATCH As String = "property data type mismatch: "

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
'+--------------------------------------------------------------------------
'| Method           | GetPropertyTypeFromVariantType
'|------------------+-------------------------------------------------------
'| Description      | Map the variant data type to a corresponding
'|                  | property data type
'|------------------+-------------------------------------------------------
'| Parameter        | propertyValue: Variant variable to map
'|------------------+-------------------------------------------------------
'| Return values    | The corresponding MsoDocProperties type
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2018-09-25  Created. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | Raises an error if mapping is not possible.
'+--------------------------------------------------------------------------
'
Private Function GetPropertyTypeFromVariantType(ByRef propertyValue As Variant) As MsoDocProperties
   Select Case VarType(propertyValue)
      Case vbInteger, vbLong, vbByte, 20
         GetPropertyTypeFromVariantType = MsoDocProperties.msoPropertyTypeNumber
      
      Case vbSingle, vbDouble, vbCurrency, vbDecimal
         GetPropertyTypeFromVariantType = MsoDocProperties.msoPropertyTypeFloat
      
      Case vbDate
         GetPropertyTypeFromVariantType = MsoDocProperties.msoPropertyTypeDate
      
      Case vbString
         GetPropertyTypeFromVariantType = MsoDocProperties.msoPropertyTypeString
      
      Case vbBoolean
         GetPropertyTypeFromVariantType = MsoDocProperties.msoPropertyTypeBoolean

      Case Else
         Err.Raise ERR_NUM_TYPE_NOT_MAPPABLE, _
                   ERR_STR_CLASS_NAME, _
                   ERR_STR_TYPE_NOT_MAPPABLE
   End Select
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
'| Parameter        | propertyWorkbook: The workbook for the property
'|                  | propertyName    : The property name
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
Public Function CustomPropertyExists(ByRef propertyWorkbook As Workbook, ByRef propertyName As String) As Boolean
   Dim aProperty As DocumentProperty

   CustomPropertyExists = False

   For Each aProperty In propertyWorkbook.CustomDocumentProperties
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
'| Parameter        | propertyWorkbook: The workbook for the property
'|                  | propertyName    : The property name
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
Public Function GetCustomProperty(ByRef propertyWorkbook As Workbook, ByRef propertyName As String) As Variant
   If CustomPropertyExists(propertyWorkbook, propertyName) Then
      GetCustomProperty = propertyWorkbook.CustomDocumentProperties.Item(propertyName).value
   Else
      Err.Raise ERR_NUM_PROPERTY_NOT_FOUND, _
                ERR_STR_CLASS_NAME, _
                ERR_STR_PROPERTY_NOT_FOUND & _
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
'| Parameter        | propertyWorkbook: The workbook for the property
'|                  | propertyName    : The property name
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
Public Function SafeGetCustomProperty(ByRef propertyWorkbook As Workbook, ByRef propertyName As String) As Variant
   If CustomPropertyExists(propertyWorkbook, propertyName) Then
      SafeGetCustomProperty = propertyWorkbook.CustomDocumentProperties.Item(propertyName).value
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
'| Parameter        | propertyWorkbook: The workbook for the property
'|                  | propertyName    : The property name
'|                  | propertyValue   : The property value
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
Public Sub SetCustomProperty(ByRef propertyWorkbook As Workbook, ByRef propertyName As String, ByVal propertyValue As Variant)
   Dim actualProperty As CustomProperty

   Dim newPropertyType As MsoDocProperties

   '
   ' Setting a property value is quite complicated. One has to check wether
   ' the property already exists, or not. One has to check wether the
   ' new type is compatible to the old type. One has to check for null
   ' or empty.
   '

   '
   ' Check, if property already exists?
   '
   If CustomPropertyExists(propertyWorkbook, propertyName) Then
      '
      ' The property already exists
      '
      ' So get it
      '
      Set actualProperty = propertyWorkbook.CustomDocumentProperties.Item(propertyName)

      If IsEmptyOrNull(propertyValue) Then
         '
         ' Delete it if the new value is empty or null
         '
         actualProperty.Delete
      Else
         '
         ' The new value is neither empty, nor null
         '
         ' Check, wether the new and the old property data type are compatible
         '
         newPropertyType = GetPropertyTypeFromVariantType(propertyValue)

         If newPropertyType = actualProperty.Type Then
            '
            ' Yes, they are. So set the new value (This is what we wanted to do in the first place)
            '
            actualProperty.value = propertyValue
         Else
            '
            ' The old and the new property data type are not compatible. Raise an error
            '
            Err.Raise ERR_NUM_TYPE_MISMATCH, _
                      ERR_STR_CLASS_NAME, _
                      ERR_STR_TYPE_MISMATCH & _
                         "Name = '" & _
                         propertyName & _
                         "', requested type = " & _
                         Format$(newPropertyType) & _
                         ", actual type = " & _
                         Format$(actualProperty.Type)
         End If
      End If
   Else
      '
      ' The property does not exist
      '
      ' Just add it (That was the easy part)
      '
      propertyWorkbook.CustomDocumentProperties.Add Name:=propertyName, _
                                                    LinkToContent:=False, _
                                                    Type:=GetPropertyTypeFromVariantType(propertyValue), _
                                                    value:=propertyValue
   End If
End Sub

'
'+--------------------------------------------------------------------------
'| Method           | DeleteProperty
'|------------------+-------------------------------------------------------
'| Description      | Delete a custom property
'|------------------+-------------------------------------------------------
'| Parameter        | propertyWorkbook: The workbook for the property
'|                  | propertyName    : The property name
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
Public Sub DeleteProperty(ByRef propertyWorkbook As Workbook, ByRef propertyName As String)
   If CustomPropertyExists(propertyWorkbook, propertyName) Then
      propertyWorkbook.CustomDocumentProperties.Item(propertyName).Delete
   Else
      Err.Raise ERR_NUM_PROPERTY_NOT_FOUND, _
                ERR_STR_CLASS_NAME, _
                ERR_STR_PROPERTY_NOT_FOUND & _
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
'| Parameter        | propertyWorkbook: The workbook for the property
'|                  | propertyName    : The property name
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
Public Sub SafeDeleteProperty(ByRef propertyWorkbook As Workbook, ByRef propertyName As String)
   If CustomPropertyExists(propertyWorkbook, propertyName) Then
      propertyWorkbook.CustomDocumentProperties.Item(propertyName).Delete
   End If
End Sub
