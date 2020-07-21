VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DAOPropertyManager"
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
'|
'|-------------------------------------------------------------------------
'| Class               | DAOPropertyManager
'|---------------------+---------------------------------------------------
'| Description         | Easily set and get DAO properties
'|---------------------+---------------------------------------------------
'| Author              | Frank Schwab
'|---------------------+---------------------------------------------------
'| Version             | 1.0.0
'|---------------------+---------------------------------------------------
'| Changes             | 2020-07-18  Created. fhs
'|---------------------+---------------------------------------------------
'
Option Compare Database
Option Explicit

Private Const ERR_PROPERTY_NOT_FOUND As Long = 3270

Public Sub SetProperty(ByRef obj As Object, _
                       ByRef propertyName As String, _
                       ByRef propertyType As Long, _
                       ByRef propertyValue As Variant)
   Dim properyList As DAO.Properties
   Dim aProperty As DAO.Property

   Set properyList = obj.Properties

   On Error Resume Next
   properyList(propertyName).Value = propertyValue

   If Err.Number <> 0 Then
      If Err.Number = ERR_PROPERTY_NOT_FOUND Then
         On Error GoTo 0

         Set aProperty = obj.CreateProperty(propertyName)

         With aProperty
            .Type = propertyType
            .Value = propertyValue
         End With

         properyList.Append aProperty

         Set aProperty = Nothing
      Else
         Err.Raise Err.Number, Err.source, Err.description, Err.HelpFile, Err.HelpContext
      End If
   End If

   Set properyList = Nothing
End Sub

Public Function GetProperty(ByRef obj As Object, _
                            ByRef propertyName As String) As Variant
   Dim properyList As DAO.Properties
   Dim aProperty As DAO.Property

   Set properyList = obj.Properties

   On Error Resume Next
   GetProperty = properyList(propertyName).Value
   
   If Err.Number <> 0 Then _
      GetProperty = Null

   On Error GoTo 0
End Function

Public Function GetPropertyWithDefault(ByRef obj As Object, _
                                       ByRef propertyName As String, _
                                       ByRef defaultValue As Variant) As Variant
   GetPropertyWithDefault = Nz(Me.GetProperty(obj, propertyName), defaultValue)
End Function
