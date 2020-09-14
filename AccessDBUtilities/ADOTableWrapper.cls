VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ADOTableWrapper"
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
'| Class               | ADOTableWrapper
'|---------------------+---------------------------------------------------
'| Description         | Easier access to ADO tables
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
Private Const TABLE_PROPERTY_HIDDEN                    As String = "Jet OLEDB:Table Hidden in Access"
Private Const TABLE_PROPERTY_LINK                      As String = "Jet OLEDB:Create Link"
Private Const TABLE_PROPERTY_LINK_CACHE_AUTHENTICATION As String = "Jet OLEDB:Cache Link Name/Password"
Private Const TABLE_PROPERTY_LINK_DATA_SOURCE          As String = "Jet OLEDB:Link Datasource"
Private Const TABLE_PROPERTY_LINK_EXCLUSIVE            As String = "Jet OLEDB:Exclusive Link"
Private Const TABLE_PROPERTY_LINK_PROVIDER_STRING      As String = "Jet OLEDB:Link Provider String"
Private Const TABLE_PROPERTY_LINK_REMOTE_TABLE_NAME    As String = "Jet OLEDB:Remote Table Name"
Private Const TABLE_PROPERTY_VALIDATION_MESSAGE        As String = "Jet OLEDB:Table Validation Text"
Private Const TABLE_PROPERTY_VALIDATION_RULE           As String = "Jet OLEDB:Table Validation Rule"

'
' Instance variable
'
Private m_Table As ADOX.Table

'
' Public attributes
'

' ADOTable
Public Property Get ADOTable() As ADOX.Table
   Set ADOTable = m_Table
End Property

Public Property Set ADOTable(ByRef newTable As ADOX.Table)
   Set m_Table = newTable
End Property


' CacheLinkAuthentication
Public Property Get CacheLinkAuthentication() As Boolean
   CacheLinkAuthentication = m_Table.Properties(TABLE_PROPERTY_LINK_CACHE_AUTHENTICATION).Value
End Property

Public Property Let CacheLinkAuthentication(ByVal newCacheLinkAuthentication As Boolean)
   m_Table.Properties(TABLE_PROPERTY_LINK_CACHE_AUTHENTICATION).Value = newCacheLinkAuthentication
End Property


' IsHidden
Public Property Get IsHidden() As Boolean
   IsHidden = m_Table.Properties(TABLE_PROPERTY_HIDDEN).Value
End Property

Public Property Let IsHidden(ByVal newIsHidden As Boolean)
   m_Table.Properties(TABLE_PROPERTY_HIDDEN).Value = newIsHidden
End Property


' LinkDataSource
Public Property Get LinkDataSource() As String
   LinkDataSource = m_Table.Properties(TABLE_PROPERTY_LINK_DATA_SOURCE).Value
End Property

Public Property Let LinkDataSource(ByRef newLinkDataSource As String)
   m_Table.Properties(TABLE_PROPERTY_LINK_DATA_SOURCE).Value = newLinkDataSource
End Property


' ExclusiveLink
Public Property Get ExclusiveLink() As Boolean
   LinkDataSource = m_Table.Properties(TABLE_PROPERTY_LINK_EXCLUSIVE).Value
End Property

Public Property Let ExclusiveLink(ByVal newExclusiveLink As Boolean)
   m_Table.Properties(TABLE_PROPERTY_LINK_EXCLUSIVE).Value = newExclusiveLink
End Property


' LinkProvider
Public Property Get LinkProvider() As String
   LinkDataSource = m_Table.Properties(TABLE_PROPERTY_LINK_PROVIDER_STRING).Value
End Property

Public Property Let LinkProvider(ByRef newLinkProvider As String)
   m_Table.Properties(TABLE_PROPERTY_LINK_PROVIDER_STRING).Value = newLinkProvider
End Property


' RemoteTableName
Public Property Get RemoteTableName() As String
   RemoteTableName = m_Table.Properties(TABLE_PROPERTY_LINK_REMOTE_TABLE_NAME).Value
End Property

Public Property Let RemoteTableName(ByRef newRemoteTableName As String)
   m_Table.Properties(TABLE_PROPERTY_LINK_REMOTE_TABLE_NAME).Value = newRemoteTableName
End Property


' ValidationMessage
Public Property Get ValidationMessage() As String
   ValidationMessage = m_Table.Properties(TABLE_PROPERTY_VALIDATION_MESSAGE).Value
End Property

Public Property Let ValidationMessage(ByRef newValidationMessage As String)
   m_Table.Properties(TABLE_PROPERTY_VALIDATION_MESSAGE).Value = newValidationMessage
End Property


' ValidationRule
Public Property Get ValidationRule() As String
   ValidationRule = m_Table.Properties(TABLE_PROPERTY_VALIDATION_RULE).Value
End Property

Public Property Let ValidationRule(ByRef newValidationRule As String)
   m_Table.Properties(TABLE_PROPERTY_VALIDATION_RULE).Value = newValidationRule
End Property


'
' Public methods
'

'
'+--------------------------------------------------------------------------
'| Method           | CreateADOTable
'|------------------+-------------------------------------------------------
'| Description      | Create this wrapper with a specific ADO table.
'|------------------+-------------------------------------------------------
'| Parameter        | tableName : Table name
'|                  | adoCatalog: ADO catalog the table should belong to
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
Public Sub CreateADOTable(ByRef tableName As String, _
                          ByRef adoCatalog As ADOX.Catalog)
   Set m_Table = New ADOX.Table

   With m_Table
      .name = tableName

      Set .parentCatalog = adoCatalog
   End With
End Sub

'
'+--------------------------------------------------------------------------
'| Method           | CreateNew
'|------------------+-------------------------------------------------------
'| Description      | Create this wrapper with a blank ADO table.
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
   Set m_Table = New ADOX.Table
End Sub

'
'+--------------------------------------------------------------------------
'| Method           | fromADOTable
'|------------------+-------------------------------------------------------
'| Description      | Create this wrapper from an existing ADO table.
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
Public Sub fromADOTable(ByRef aTable As ADOX.Table)
   Set m_Table = aTable
End Sub
