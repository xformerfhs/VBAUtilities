VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DBTableLinkHelper"
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
'| Class               | DBTableLinkHelper
'|---------------------+---------------------------------------------------
'| Description         | Get and change the path to a linked table
'|---------------------+---------------------------------------------------
'| Author              | Frank Schwab
'|---------------------+---------------------------------------------------
'| Version             | 1.0.0
'|---------------------+---------------------------------------------------
'| Changes             | 2020-07-20  Created. fhs
'|---------------------+---------------------------------------------------
'

Option Compare Database
Option Explicit

'
' Private methods
'
Private Function GetDBPathFromTablePath(ByRef dbPath As String) As String
   Dim pos As Long

   pos = InStr(1, dbPath, "=")
   
   If pos <> 0 Then
      GetDBPathFromTablePath = Right$(dbPath, Len(dbPath) - pos)
   Else
      GetDBPathFromTablePath = CurrentProject.Path
   End If
End Function

'
' Public methods
'
Public Function GetDBPath(ByRef tableName As String) As String
   Dim db As DAO.Database

   Set db = DBEngine.Workspaces(0).Databases(0)
   
   Dim tdf As DAO.TableDef
   
   Set tdf = db.TableDefs(tableName)

   GetDBPath = GetDBPathFromTablePath(tdf.Connect)
   
   Set tdf = Nothing
   
   Set db = Nothing
End Function

Public Function ChangeDBPath(ByRef fromDBPath As String, ByRef toDBPath As String) As Integer
   Dim count As Integer
   
   Dim db As DAO.Database

   Set db = DBEngine.Workspaces(0).Databases(0)
   
   Dim tdf As DAO.TableDef
   
   For Each tdf In db.TableDefs
      If Len(tdf.Connect) <> 0 Then
         If GetDBPathFromTablePath(tdf.Connect) = fromDBPath Then
            tdf.Connect = ";DATABASE=" & toDBPath
            count = count + 1
         End If
      End If
   Next

   ChangeDBPath = count

   db.TableDefs.Refresh

   Set db = Nothing
End Function
