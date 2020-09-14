VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ADOFieldWrapper"
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
'| Class               | ADOFieldWrapper
'|---------------------+---------------------------------------------------
'| Description         | Easier access to ADO fields
'|---------------------+---------------------------------------------------
'| Author              | Frank Schwab
'|---------------------+---------------------------------------------------
'| Version             | 1.0.1
'|---------------------+---------------------------------------------------
'| Changes             | 2020-07-18  Created. fhs
'|                     | 2020-09-14  Corrected comments. fhs
'|---------------------+---------------------------------------------------
'

Option Explicit

'
' Private constants
'
Private Const FIELD_PROPERTY_AUTO_VALUE          As String = "Autoincrement"
Private Const FIELD_PROPERTY_DESCRIPTION         As String = "Description"
Private Const FIELD_PROPERTY_NULLABLE            As String = "Nullable"
Private Const FIELD_PROPERTY_ALLOW_0_LENGTH      As String = "Jet OLEDB:Allow Zero Length"
Private Const FIELD_PROPERTY_UNICODE_COMPRESSION As String = "Jet OLEDB:Compressed UNICODE Strings"
Private Const FIELD_PROPERTY_VALIDATION_MESSAGE  As String = "Jet OLEDB:Column Validation Text"
Private Const FIELD_PROPERTY_VALIDATION_RULE     As String = "Jet OLEDB:Column Validation Rule"
Private Const FIELD_PROPERTY_DEFAULT_VALUE       As String = "Default"

Private Const INVALID_SIZE As Long = -1

'
' Instance variable
'
Private m_Field As ADOX.Column

'
' Public attributes
'

'
' ADOField
Public Property Get ADOField() As ADOX.Column
   Set ADOField = m_Field
End Property

Public Property Set ADOField(ByRef newField As ADOX.Column)
   Set m_Field = newField
End Property


' AllowZeroLength
Public Property Get AllowZeroLength() As Boolean
   AllowZeroLength = m_Field.Properties(FIELD_PROPERTY_ALLOW_0_LENGTH).Value
End Property

Public Property Let AllowZeroLength(ByVal newAllowZeroLength As Boolean)
   m_Field.Properties(FIELD_PROPERTY_ALLOW_0_LENGTH).Value = newAllowZeroLength
End Property


' DefaultValue
Public Property Get DefaultValue() As Variant
   Required = m_Field.Properties(FIELD_PROPERTY_DEFAULT_VALUE).Value
End Property

Public Property Let DefaultValue(ByVal newDefaultValue As Variant)
   m_Field.Properties(FIELD_PROPERTY_DEFAULT_VALUE).Value = newDefaultValue
End Property


' Description
Public Property Get Description() As String
   Description = m_Field.Properties(FIELD_PROPERTY_DESCRIPTION).Value
End Property

Public Property Let Description(ByRef newDescription As String)
   m_Field.Properties(FIELD_PROPERTY_DESCRIPTION).Value = newDescription
End Property


' IsAutoValue
Public Property Get IsAutoValue() As Boolean
   IsAutoValue = m_Field.Properties(FIELD_PROPERTY_AUTO_VALUE).Value
End Property

Public Property Let IsAutoValue(ByVal newIsAutoValue As Boolean)
   m_Field.Properties(FIELD_PROPERTY_AUTO_VALUE).Value = newIsAutoValue
End Property


' Required
Public Property Get Required() As Boolean
   Required = Not m_Field.Properties(FIELD_PROPERTY_NULLABLE).Value
End Property

Public Property Let Required(ByVal newRequired As Boolean)
   m_Field.Properties(FIELD_PROPERTY_NULLABLE).Value = Not newRequired
End Property


' ValidationMessage
Public Property Get ValidationMessage() As String
   ValidationMessage = m_Field.Properties(FIELD_PROPERTY_VALIDATION_MESSAGE).Value
End Property

Public Property Let ValidationMessage(ByRef newValidationMessage As String)
   m_Field.Properties(FIELD_PROPERTY_VALIDATION_MESSAGE).Value = newValidationMessage
End Property


' ValidationRule
Public Property Get ValidationRule() As String
   ValidationRule = m_Field.Properties(FIELD_PROPERTY_VALIDATION_RULE).Value
End Property

Public Property Let ValidationRule(ByRef newValidationRule As String)
   m_Field.Properties(FIELD_PROPERTY_VALIDATION_RULE).Value = newValidationRule
End Property


' UnicodeCompression
Public Property Get UnicodeCompression() As Boolean
   UnicodeCompression = m_Field.Properties(FIELD_PROPERTY_UNICODE_COMPRESSION).Value
End Property

Public Property Let UnicodeCompression(ByVal newUnicodeCompression As Boolean)
   m_Field.Properties(FIELD_PROPERTY_UNICODE_COMPRESSION).Value = newUnicodeCompression
End Property


'
' Public methods
'

'
'+--------------------------------------------------------------------------
'| Method           | CreateADOField
'|------------------+-------------------------------------------------------
'| Description      | Create this wrapper with a specific ADO field.
'|------------------+-------------------------------------------------------
'| Parameter        | fieldName : Field name
'|                  | fieldtype : Field type
'|                  | adoCatalog: ADO catalog the field should belong to
'|                  | fieldSize : Size of the field, if it is necessary
'|                  |             for the field definition (e.g. strings)
'|------------------+-------------------------------------------------------
'| Return values    | ./.
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2020-07-18  Created. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | ./.
'+--------------------------------------------------------------------------
'
Public Sub CreateADOField(ByRef fieldName As String, _
                          ByVal fieldType As ADOX.DataTypeEnum, _
                          ByRef adoCatalog As ADOX.Catalog, _
                          Optional ByVal fieldSize As Long = INVALID_SIZE)
   Set m_Field = New ADOX.Column

   With m_Field
      .name = fieldName
      .Type = fieldType

      Set .parentCatalog = adoCatalog

      If fieldSize <> INVALID_SIZE Then
         .DefinedSize = fieldSize
      End If
   End With
End Sub

'
'+--------------------------------------------------------------------------
'| Method           | CreateNew
'|------------------+-------------------------------------------------------
'| Description      | Create this wrapper with a blank ADO field.
'|------------------+-------------------------------------------------------
'| Parameter        | ./.
'|------------------+-------------------------------------------------------
'| Return values    | ./.
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2020-07-18  Created. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | ./.
'+--------------------------------------------------------------------------
'
Public Sub CreateNew()
   Set m_Field = New ADOX.Column
End Sub

'
'+--------------------------------------------------------------------------
'| Method           | FromADOField
'|------------------+-------------------------------------------------------
'| Description      | Create this wrapper from an existing ADO field.
'|------------------+-------------------------------------------------------
'| Parameter        | ./.
'|------------------+-------------------------------------------------------
'| Return values    | ./.
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2020-07-18  Created. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | ./.
'+--------------------------------------------------------------------------
'
Public Sub FromADOField(ByRef aField As ADOX.Column)
   Set m_Field = aField
End Sub

'
'+--------------------------------------------------------------------------
'| Method           | SetAutoValue
'|------------------+-------------------------------------------------------
'| Description      | Set the AutoValue attribute of the field in this wrapper.
'|------------------+-------------------------------------------------------
'| Parameter        | ./.
'|------------------+-------------------------------------------------------
'| Return values    | ./.
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2020-07-18  Created. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | ./.
'+--------------------------------------------------------------------------
'
Public Sub SetAutoValue()
   Me.IsAutoValue = True
End Sub

'
'+--------------------------------------------------------------------------
'| Method           | NoAutoValue
'|------------------+-------------------------------------------------------
'| Description      | Clear the AutoValue attribute of the field in this wrapper.
'|------------------+-------------------------------------------------------
'| Parameter        | ./.
'|------------------+-------------------------------------------------------
'| Return values    | ./.
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2020-07-18  Created. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | ./.
'+--------------------------------------------------------------------------
'
Public Sub NoAutoValue()
   Me.IsAutoValue = False
End Sub
