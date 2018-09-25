VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WorkbookCustomPropertyHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'''
''' Wrapper class for managing Excel workbook custom properties, as the MS supplied API
''' is cumbersome and incomplete.
'''

'
' Error constants
'
Private Const MODULE_NAME As String = "WorkbookCustomPropertyHandler"
Private Const ERROR_NUMBER_BASE As Long = 38803

Private Const ERROR_NUMBER_NOT_FOUND As Long = ERROR_NUMBER_BASE + 0
Private Const ERROR_TEXT_NOT_FOUND As String = "There is no property named '"

Private Const ERROR_NUMBER_TYPE_NOT_MAPPABLE As Long = ERROR_NUMBER_BASE + 1
Private Const ERROR_TEXT_TYPE_NOT_MAPPABLE As String = "variant data type can not be mapped to document property data type"

Private Const ERROR_NUMBER_TYPE_MISMATCH As Long = ERROR_NUMBER_BASE + 2
Private Const ERROR_TEXT_TYPE_MISMATCH As String = "property data type mismatch: "

'
' Helper constant for StrComp
'
Private Const STRCOMP_EQUAL As Integer = 0


'
' Private methods
'

'''
''' Check wether the variant data type is empty or null
'''
Private Function IsEmptyOrNull(ByRef propertyValue As Variant) As Boolean
   IsEmptyOrNull = (VarType(propertyValue) <= 1)
End Function
   
'''
''' Map the variant data type to a corresponding property data type
'''
''' Raise an error if such a mapping is not possible
'''
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
         Err.Raise ERROR_NUMBER_TYPE_NOT_MAPPABLE, _
                   MODULE_NAME, _
                   ERROR_TEXT_TYPE_NOT_MAPPABLE
   End Select
End Function


'
' Public methods
'

'''
''' Check wether a property with a given name exists
'''
''' Property name comparisons are case-insensitive! This is the way Excel works.
''' A property with the name "aProperty" is the same as a property with the name "ApRoPeRtY".
'''
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

'''
''' Get the value of a property and raise an error if the property does not exist
'''
Public Function GetCustomProperty(ByRef propertyWorkbook As Workbook, ByRef propertyName As String) As Variant
   If CustomPropertyExists(propertyWorkbook, propertyName) Then
      GetCustomProperty = propertyWorkbook.CustomDocumentProperties.Item(propertyName).value
   Else
      Err.Raise ERROR_NUMBER_NOT_FOUND, _
                MODULE_NAME, _
                ERROR_TEXT_NOT_FOUND & _
                   propertyName & _
                   "'"
   End If
End Function

'''
''' Get the value of a property and do not raise an error if the property does not exist.
''' Instead return a value of "Empty".
'''
Public Function SafeGetCustomProperty(ByRef propertyWorkbook As Workbook, ByRef propertyName As String) As Variant
   If CustomPropertyExists(propertyWorkbook, propertyName) Then
      SafeGetCustomProperty = propertyWorkbook.CustomDocumentProperties.Item(propertyName).value
   Else
      SafeGetCustomProperty = Empty
   End If
End Function

'''
''' Set a property to a value
'''
''' The property data type is derived from the variant data type.
'''
''' This method raises an error if the variant data type can not be matched to a property data type
''' and if the property already exists and the new and the old property data types do not match.
'''
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
            ' The old and the new property data type are not compatible. raise an error
            '
            Err.Raise ERROR_NUMBER_TYPE_MISMATCH, _
                      MODULE_NAME, _
                      ERROR_TEXT_TYPE_MISMATCH & _
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
      ' Just add it (That was easy)
      '
      propertyWorkbook.CustomDocumentProperties.Add Name:=propertyName, _
                                                    LinkToContent:=False, _
                                                    Type:=GetPropertyTypeFromVariantType(propertyValue), _
                                                    value:=propertyValue
   End If
End Sub

'''
''' Delete a property and raise an error if it did not exist
'''
Public Sub DeleteProperty(ByRef propertyWorkbook As Workbook, ByRef propertyName As String)
   If CustomPropertyExists(propertyWorkbook, propertyName) Then
      propertyWorkbook.CustomDocumentProperties.Item(propertyName).Delete
   Else
      Err.Raise ERROR_NUMBER_NOT_FOUND, _
                MODULE_NAME, _
                ERROR_TEXT_NOT_FOUND & _
                   propertyName & _
                   "'"
   End If
End Sub

'''
''' Delete a property without throwing an error if it did not exist
'''
Public Sub SafeDeleteProperty(ByRef propertyWorkbook As Workbook, ByRef propertyName As String)
   If CustomPropertyExists(propertyWorkbook, propertyName) Then
      propertyWorkbook.CustomDocumentProperties.Item(propertyName).Delete
   End If
End Sub
