VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MessageManager"
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
'| Class               | MessageManager
'|---------------------+---------------------------------------------------
'| Description         | Get Windows API error messages and format
'|                     | messages in general.
'|---------------------+---------------------------------------------------
'| Author              | Frank Schwab
'|---------------------+---------------------------------------------------
'| Version             | 2.1.1
'|---------------------+---------------------------------------------------
'| Changes             | 2002-03-15  Created. fhs
'|                     | 2006-03-06  Added formatting of general messages. fhs
'|                     | 2014-03-31  V2: Simplified, added Windows NT status
'|                     |                 codes. fhs
'|                     | 2014-04-01  Simplified Windows NT status code handling. fhs
'|                     | 2015-01-06  Added module error handling. fhs
'|                     | 2020-08-24  Implemented 64 bit compatibility. fhs
'+-------------------------------------------------------------------------

Option Explicit

'
' Private constants
'

Private Const MESSAGE_BUFFER_SIZE As Long = 2048

'
' Windows API constants
'
Private Const FORMAT_MESSAGE_IGNORE_INSERTS As Long = &H200&
Private Const FORMAT_MESSAGE_FROM_STRING    As Long = &H400&
Private Const FORMAT_MESSAGE_FROM_HMODULE   As Long = &H800&
Private Const FORMAT_MESSAGE_FROM_SYSTEM    As Long = &H1000&
Private Const FORMAT_MESSAGE_ARGUMENT_ARRAY As Long = &H2000&
Private Const FORMAT_MESSAGE_MAX_WIDTH_MASK As Long = &HFF&

'
' Windows API declarations
'
Private Declare PtrSafe Function FormatMessage _
  Lib "kernel32.dll" _
  Alias "FormatMessageW" ( _
  ByVal dwFlags As Long, _
  ByVal lpSource As LongPtr, _
  ByVal dwMessageId As Long, _
  ByVal dwLanguageId As Long, _
  ByVal lpBuffer As LongPtr, _
  ByVal nSize As Long, _
  ByRef Arguments As Long _
  ) As Long

Private Declare PtrSafe Function RtlNtStatusToDosError _
   Lib "ntdll.dll" ( _
   ByVal dwStatus As Long _
   ) As Long

Private Declare PtrSafe Function GetModuleHandle _
   Lib "kernel32.dll" _
   Alias "GetModuleHandleA" ( _
   ByVal lpModuleName As String _
   ) As Long
   
   
'
' Public methods
'

'
'+--------------------------------------------------------------------------
'| Method           | GetMessageForWindowsErrorCode
'|------------------+-------------------------------------------------------
'| Purpose          | Get error text that belongs to an error code return
'|                  | by a Windows API call
'|------------------+-------------------------------------------------------
'| Parameter        | windowsErrorCode: Error code from Err.LastDllError
'|------------------+-------------------------------------------------------
'| Return value     | Text belonging to the error code or a message with
'|                  | the error code and the text "unknown error code".
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2002-03-15  Created. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | This Method is based on code from the Microsoft
'|                  | Developer Network that can be found here:
'|                  | http://www.microsoft.com/germany/msdn/library/visualtools/vb6/APIFehlermeldungenImKlartext.mspx
'|------------------+-------------------------------------------------------
'| Typical call     | rc = SomeWindowsAPIFunction(...)
'|                  | If rc = RC_FAILED Then
'|                  |    rc = Err.LastDllError
'|                  |    Err.Raise M_N_ERROR_NUM, M_STR_ERROR_SOURCE , _
'|                  |              "SomeWindowsAPIFunction returned error code 0x" & Hex$(rc) & ": " & _
'|                  |              mm.GetMessageForWindowsErrorCode(rc))
'|                  | Else
'|                  |    ' Everything worked fine
'|                  | End If
'+--------------------------------------------------------------------------
Public Function GetMessageForWindowsErrorCode(ByVal windowsErrorCode As Long) As String
   Dim messageBuffer As String  ' Buffer for error message
   Dim bufferLength  As Long    ' Size of error message

' Initialize the buffer that will receive the message text
   messageBuffer = Space$(MESSAGE_BUFFER_SIZE)

' Get the message text from Windows
   bufferLength = FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM Or _
                                FORMAT_MESSAGE_MAX_WIDTH_MASK Or _
                                FORMAT_MESSAGE_IGNORE_INSERTS, _
                                0&, _
                                windowsErrorCode, _
                                0&, _
                                StrPtr(messageBuffer), _
                                Len(messageBuffer), _
                                0&)

   If bufferLength > 0 Then
      ' Error code is known to Windows. Return the message with the correct length.
      GetMessageForWindowsErrorCode = Left$(messageBuffer, bufferLength)
   Else
      ' Error code is unknown to Windows. Return a default error message.
      GetMessageForWindowsErrorCode = "Unknown error code: " & _
                                      Format$(windowsErrorCode) & _
                                      " (0x" & _
                                      Hex$(windowsErrorCode) & _
                                      ")"
   End If
End Function

'
'+--------------------------------------------------------------------------
'| Method           | GetMessageForModuleErrorCode
'|------------------+-------------------------------------------------------
'| Purpose          | Get error text that belongs to an error code return
'|                  | by a module API call
'|------------------+-------------------------------------------------------
'| Parameter        | moduleErrorCode: Error code from Err.LastDllError
'|                  | moduleName: Name of the dll that set the error code
'|------------------+-------------------------------------------------------
'| Return value     | Text belonging to the error code or a message with
'|                  | the error code and the text "unknown error code".
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2015-01-06  Created. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | ./.
'|------------------+-------------------------------------------------------
'| Typical call     | rc = SomeModuleAPIFunction(...)
'|                  | If rc = RC_FAILED Then
'|                  |    rc = Err.LastDllError
'|                  |    Err.Raise M_N_ERROR_NUM, M_STR_ERROR_SOURCE , _
'|                  |              "SomeModuleAPIFunction returned error code 0x" & Hex$(rc) & ": " & _
'|                  |              mm.GetMessageForModuleErrorCode(rc, moduleName))
'|                  | Else
'|                  |    ' Everything worked fine
'|                  | End If
'+--------------------------------------------------------------------------
Public Function GetMessageForModuleErrorCode(ByVal moduleErrorCode As Long, _
                                             ByRef moduleName As String) As String
   Dim messageBuffer As String  ' Buffer for error message
   Dim bufferLength  As Long    ' Size of error message
   Dim moduleHandle  As Long    ' Handle for module
   Dim loadError     As Long    ' Module load error code

   moduleHandle = GetModuleHandle(moduleName)

   If moduleHandle <> 0& Then
      ' Initialize the buffer that will receive the message text
      messageBuffer = Space$(MESSAGE_BUFFER_SIZE)

      ' Get the message text from the module
      bufferLength = FormatMessage(FORMAT_MESSAGE_FROM_HMODULE Or _
                                   FORMAT_MESSAGE_MAX_WIDTH_MASK Or _
                                   FORMAT_MESSAGE_IGNORE_INSERTS, _
                                   moduleHandle, _
                                   moduleErrorCode, _
                                   0&, _
                                   StrPtr(messageBuffer), _
                                   Len(messageBuffer), _
                                   0&)
   
      If bufferLength > 0 Then
         ' Error code is known to the module. Return the message with the correct length.
         GetMessageForModuleErrorCode = Left$(messageBuffer, bufferLength)
      Else
         ' Error code is unknown to the module. Try system error message.
         GetMessageForModuleErrorCode = Me.GetMessageForWindowsErrorCode(moduleErrorCode)
      End If
   Else
      loadError = Err.LastDllError

      GetMessageForModuleErrorCode = "Could not load module '" & _
                                     moduleName & _
                                     "' to get message for error code " & _
                                        Format$(moduleErrorCode) & _
                                        " (0x" & _
                                        Hex$(moduleErrorCode) & _
                                        "). Load error code is " & _
                                        Format$(loadError) & _
                                        " (0x" & _
                                        Hex$(loadError) & _
                                        ")"
   End If
End Function

'
'+--------------------------------------------------------------------------
'| Method           | GetWindowsErrorCodeFromNTStatusCode
'|------------------+-------------------------------------------------------
'| Purpose          | Get error error code that correpsonds to an NT
'|                  | status code returned by a Windows API call
'|------------------+-------------------------------------------------------
'| Parameter        | windowsErrorCode: Error code returned from API call
'|                  |                   returning an NT status code.
'|------------------+-------------------------------------------------------
'| Return value     | Windows error code corresponding to the NT status code.
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2014-04-01  Created. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | ./.
'|------------------+-------------------------------------------------------
'| Typical call     | rc = SomeNTSCReturningFunction(...)
'|                  | If rc <> STATUS_SUCCESS Then
'|                  |    Err.Raise M_N_ERROR_NUM, M_STR_ERROR_SOURCE , _
'|                  |              "SomeNTSCReturningFunction returned error code  0x" & Hex$(rc) & ": " & _
'|                  |              mm.GetMessageForWindowsErrorCode(mm.GetWindowsErrorCodeFromNTStatusCode(rc)))
'|                  | Else
'|                  |    ' Everything worked fine
'|                  | End If
'+--------------------------------------------------------------------------
Public Function GetWindowsErrorCodeFromNTStatusCode(ByVal ntStatusCode As Long) As Long
   GetWindowsErrorCodeFromNTStatusCode = RtlNtStatusToDosError(ntStatusCode)
End Function

'
'+--------------------------------------------------------------------------
'| Method           | GetMessageForNTStatusCode
'|------------------+-------------------------------------------------------
'| Purpose          | Get error text that belongs to an NT status code
'|                  | returned by a Windows API call
'|------------------+-------------------------------------------------------
'| Parameter        | status code: Status code returned from API call.
'|------------------+-------------------------------------------------------
'| Return value     | Text belonging to the status code or a message with
'|                  | the status code and the text "unknown error code".
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2014-03-31  Created. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | ./.
'|------------------+-------------------------------------------------------
'| Typical call     | rc = SomeNTSCReturningFunction(...)
'|                  | If rc <> STATUS_SUCCESS Then
'|                  |    Err.Raise M_N_ERROR_NUM, M_STR_ERROR_SOURCE , _
'|                  |              "SomeNTSCReturningFunction returned error code 0x" & Hex$(rc) & ": " & _
'|                  |              mm.GetMessageForNTStatusCode(rc))
'|                  | Else
'|                  |    ' Everything worked fine
'|                  | End If
'+--------------------------------------------------------------------------
Public Function GetMessageForNTStatusCode(ByVal ntStatusCode As Long) As String
   GetMessageForNTStatusCode = Me.GetMessageForWindowsErrorCode( _
                                  Me.GetWindowsErrorCodeFromNTStatusCode(ntStatusCode))
End Function

'+--------------------------------------------------------------------------
'| Method           | FormatMessageWithParameters
'|------------------+-------------------------------------------------------
'| Purpose          | Create a message from a format string and parameters
'|                  | that are to substituted for the placeholders.
'|------------------+-------------------------------------------------------
'| Parameters       | formatText:    Formattierungs-Zeichenkette
'|                  | parameterList: Parameter, die eingefügt werden
'|------------------+-------------------------------------------------------
'| Returns          | Formatted string
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2006-03-06  Created. fhs
'|------------------+-------------------------------------------------------
'| Typical call     | aMessage = mm.formatMessageWithParameters("Calling %1 returned
'|                  |              error code %2", "CryptContextCreate", rc)
'|------------------+-------------------------------------------------------
'| Remarks          | Parameters begin are marked by a "%" followed by one
'|                  | digit. The digit indicates which element of the
'|                  | parameter list should be substituted.
'|                  |
'|                  | Found in the microsoft.public.vb.winapi message group
'|                  | Original author: Tom Esh
'|                  | Clarified and simplified by Frank Schwab
'+--------------------------------------------------------------------------
Public Function FormatMessageWithParameters(ByRef formatText As String, _
                                            ParamArray parameterList() As Variant) As String
   Dim messageBuffer As String
   Dim parameterAsString() As String
   Dim parameterPointer() As LongPtr
   Dim i As Integer
   Dim parameterListLowerBound As Integer
   Dim parameterListUpperBound As Integer
   Dim bufferSize As Long
   Dim messageLength As Long

   bufferSize = Len(formatText) + 1 ' Accomodate for the trailing '\0' byte

   parameterListLowerBound = LBound(parameterList)
   parameterListUpperBound = UBound(parameterList)

   '
   ' Copy all parameter strings into two parameter arrays
   ' One gets the parameter values as strings and the other
   ' gets the addresses of these strings
   '
   ReDim parameterPointer(parameterListLowerBound To parameterListUpperBound)
   ReDim parameterAsString(parameterListLowerBound To parameterListUpperBound)

   For i = parameterListLowerBound To parameterListUpperBound
       parameterAsString(i) = CStr(parameterList(i))
       parameterPointer(i) = StrPtr(parameterAsString(i))
       bufferSize = bufferSize + Len(parameterList(i))
   Next i

   '
   ' Get the formatted message
   '
   messageBuffer = String$(bufferSize, vbNullChar)

   messageLength = FormatMessage(FORMAT_MESSAGE_FROM_STRING Or _
                                 FORMAT_MESSAGE_ARGUMENT_ARRAY, _
                                 StrPtr(formatText), _
                                 0&, _
                                 0&, _
                                 StrPtr(messageBuffer), _
                                 bufferSize, _
                                 parameterPointer(parameterListLowerBound))

   '
   ' Return with the correct length
   '
   FormatMessageWithParameters = Left$(messageBuffer, messageLength)
End Function
