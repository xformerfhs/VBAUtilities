VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StatusLine"
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
'| Class               | StatusLine
'|---------------------+---------------------------------------------------
'| Description         | Show and clear Access status line
'|---------------------+---------------------------------------------------
'| Author              | Frank Schwab
'|---------------------+---------------------------------------------------
'| Version             | 1.0.0
'|---------------------+---------------------------------------------------
'| Changes             | 2018-09-25  Created. fhs
'|---------------------+---------------------------------------------------
'| Remarks             | The status area is automatically cleared upon
'|                     | termination of this instance.
'|-------------------------------------------------------------------------
'

Option Compare Database
Option Explicit

'
' Public methods
'

'
'+--------------------------------------------------------------------------
'| Method           | Clear
'|------------------+-------------------------------------------------------
'| Description      | Clear the status area
'|------------------+-------------------------------------------------------
'| Parameter        | ./.
'|------------------+-------------------------------------------------------
'| Return values    | ./.
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2018-09-25  Created. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | ./.
'+--------------------------------------------------------------------------
'
Public Sub Clear()
   SysCmd acSysCmdClearStatus
   
   DoEvents	  ' This is necessary for the status area to be cleared
End Sub

'
'+--------------------------------------------------------------------------
'| Method           | Show
'|------------------+-------------------------------------------------------
'| Description      | Show a text in the status area
'|------------------+-------------------------------------------------------
'| Parameter        | aText: Text to show
'|------------------+-------------------------------------------------------
'| Return values    | ./.
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2018-09-25  Created. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | ./.
'+--------------------------------------------------------------------------
'
Public Sub Show(ByRef aText As String)
   If Len(aText) > 0 Then
      SysCmd acSysCmdSetStatus, aText
   
      DoEvents	' This is necessary for the text to be shown
   Else
      Me.Clear
   End If
End Sub

'
' Class methods
'
Private Sub Class_Terminate()
   Me.Clear	  ' Clear the status area upon termination of this class (very handy :-)
End Sub
