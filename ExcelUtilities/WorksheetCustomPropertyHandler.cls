VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WorksheetCustomPropertyHandler"
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
Private Const MODULE_NAME As String = "WorksheetCustomPropertyHandler"
Private Const ERROR_NUMBER_BASE As Long = 27792

Private Const ERROR_NUMBER_NOT_FOUND As Long = ERROR_NUMBER_BASE + 0
Private Const ERROR_TEXT_NOT_FOUND As String = "There is no property named '"


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
   

'
' Public methods
'

'''
''' Check wether a property with a given name exists
'''
''' Property name comparisons are case-insensitive! This is the way Excel works.
''' A property with the name "aProperty" is the same as a property with the name "ApRoPeRtY".
'''
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

'''
''' Get the value of a property and raise an error if the property does not exist
'''
Public Function GetCustomProperty(ByRef propertyWorksheet As Worksheet, ByRef propertyName As String) As Variant
   If CustomPropertyExists(propertyWorksheet, propertyName) Then
      GetCustomProperty = propertyWorksheet.CustomProperties.Item(propertyName).value
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
Public Function SafeGetCustomProperty(ByRef propertyWorksheet As Worksheet, ByRef propertyName As String) As Variant
   If CustomPropertyExists(propertyWorksheet, propertyName) Then
      SafeGetCustomProperty = propertyWorksheet.CustomProperties.Item(propertyName).value
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
      ' Just add it (That was easy)
      '
      propertyWorksheet.CustomProperties.Add Name:=propertyName, value:=propertyValue
   End If
End Sub

'''
''' Delete a property and raise an error if it did not exist
'''
Public Sub DeleteProperty(ByRef propertyWorksheet As Worksheet, ByRef propertyName As String)
   If CustomPropertyExists(propertyWorksheet, propertyName) Then
      propertyWorksheet.CustomProperties.Item(propertyName).Delete
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
Public Sub SafeDeleteProperty(ByRef propertyWorksheet As Worksheet, ByRef propertyName As String)
   If CustomPropertyExists(propertyWorksheet, propertyName) Then
      propertyWorksheet.CustomProperties.Item(propertyName).Delete
   End If
End Sub
