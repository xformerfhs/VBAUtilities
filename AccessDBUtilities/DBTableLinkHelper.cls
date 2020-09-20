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
'|
'|-------------------------------------------------------------------------
'| Class               | DBTableLinkHelper
'|---------------------+---------------------------------------------------
'| Description         | Get and change the path to a linked table
'|---------------------+---------------------------------------------------
'| Author              | Frank Schwab
'|---------------------+---------------------------------------------------
'| Version             | 1.0.1
'|---------------------+---------------------------------------------------
'| Changes             | 2018-09-25  Created. fhs
'|                     | 2020-09-20  Made get of current DB instance robust. fhs
'|---------------------+---------------------------------------------------
'

Option Compare Database
Option Explicit

'
' Private methods
'

'
'+--------------------------------------------------------------------------
'| Method           | GetCurrentDBPath
'|------------------+-------------------------------------------------------
'| Description      | Get the file path of the current database
'|------------------+-------------------------------------------------------
'| Parameter        | ./.
'|------------------+-------------------------------------------------------
'| Return values    | Current database path
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2020-09-20  Created. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | ./.
'+--------------------------------------------------------------------------
'
Private Function GetCurrentDBPath() As String
   Dim cdb As DAO.Database
   
   Set cdb = CurrentDb
   
   GetCurrentDBPath = cdb.Name
   
   cdb.Close
   
   Set cdb = Nothing
End Function

'
'+--------------------------------------------------------------------------
'| Method           | GetThisDBInstance
'|------------------+-------------------------------------------------------
'| Description      | Get the instance of the current database
'|------------------+-------------------------------------------------------
'| Parameter        | ./.
'|------------------+-------------------------------------------------------
'| Return values    | Database instance of the current database
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2020-09-20  Created. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | This method is like the "CurrentDB" function only
'|                  | that it returns the instance of the current database
'|                  | and not a copy of it.
'+--------------------------------------------------------------------------
'
Private Function GetThisDBInstance() As DAO.Database
   Dim currentDBPath As String
   
   currentDBPath = GetCurrentDBPath
   
   Dim ws As DAO.Workspace
   Dim db As DAO.Database

   For Each ws In DBEngine.Workspaces
      For Each db In ws.Databases
         If db.Name = currentDBPath Then
            Set GetThisDBInstance = db
            Exit Function
         End If
      Next
   Next
End Function

'
'+--------------------------------------------------------------------------
'| Method           | GetDBPathFromTablePath
'|------------------+-------------------------------------------------------
'| Description      | Extract database path from table path
'|------------------+-------------------------------------------------------
'| Parameter        | tablePath: Path of table
'|------------------+-------------------------------------------------------
'| Return values    | Database path
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2018-09-25  Created. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | ./.
'+--------------------------------------------------------------------------
'
Private Function GetDBPathFromTablePath(ByRef tablePath As String) As String
   Dim pos As Long

   pos = InStr(1, tablePath, "=")
   
   If pos <> 0 Then
      GetDBPathFromTablePath = Right$(tablePath, Len(tablePath) - pos)
   Else
      GetDBPathFromTablePath = CurrentProject.Path
   End If
End Function

'
' Public methods
'

'
'+--------------------------------------------------------------------------
'| Method           | GetDBPath
'|------------------+-------------------------------------------------------
'| Description      | Get database path for a linked table
'|------------------+-------------------------------------------------------
'| Parameter        | tableName: Table name
'|------------------+-------------------------------------------------------
'| Return values    | Database path
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2018-09-25  Created. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | ./.
'+--------------------------------------------------------------------------
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

'
'+--------------------------------------------------------------------------
'| Method           | ChangeDBPath
'|------------------+-------------------------------------------------------
'| Description      | Change the database path for all linked tables
'|                  | from one path to another
'|------------------+-------------------------------------------------------
'| Parameter        | fromDBPath: Database path to change
'|                  | toDBPath  : New database path
'|------------------+-------------------------------------------------------
'| Return values    | Number of paths changed
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2018-09-25  Created. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | ./.
'+--------------------------------------------------------------------------
'
Public Function ChangeDBPath(ByRef fromDBPath As String, ByRef toDBPath As String) As Integer
   Dim count As Integer
   
   Dim db As DAO.Database

   Set db = GetCurrentDBPath   ' This needs to be the exact instance that is currently running. Do not use CurrentDB

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

   Set db = Nothing  ' Never do close the current DB instance!
End Function
