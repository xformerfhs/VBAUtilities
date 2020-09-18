VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SetPriorityClass"
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
'| Class               | Set priority class
'|---------------------+---------------------------------------------------
'| Description         | Set the priority class of the currently running program
'|---------------------+---------------------------------------------------
'| Author              | Frank Schwab
'|---------------------+---------------------------------------------------
'| Version             | 1.0.1
'|---------------------+---------------------------------------------------
'| Changes             | 2020-07-20  Created. fhs
'|                     | 2020-08-24  Made 64 bit compatible. fhs
'|---------------------+---------------------------------------------------
'

Option Explicit

'
' Error constants
'
Private Const CLASS_NAME As String = "SetPriorityClass"

Private Const ERROR_NUMBER_START As Long = 46658

Private Const ERROR_NUMBER_INVALID_PRIORITY_CLASS As Long = ERROR_NUMBER_START + 0
Private Const ERROR_TEXT_INVALID_PRIORITY_CLASS   As String = "Invalid priority class value: "

Private Const ERROR_NUMBER_CAN_NOT_SET_PRIORITY_CLASS As Long = ERROR_NUMBER_START + 1
Private Const ERROR_NUMBER_CAN_NOT_GET_PRIORITY_CLASS As Long = ERROR_NUMBER_START + 2
Private Const ERROR_MESSAGE_FOR_API As String = "Could not %1 priority class. Error number: %2. Error message: '%3'"

'
' Windows API constants
'
Private Const ERROR_FAILURE As Long = 0

Private Const PROCESS_QUERY_INFORMATION As Long = &H400
Private Const PROCESS_SET_INFORMATION   As Long = &H200

' Win2000: Indicates a process that has priority above
' NORMAL_PRIORITY_CLASS but below HIGH_PRIORITY_CLASS.
Private Const ABOVE_NORMAL_PRIORITY_CLASS As Long = &H8000

' Win2000: Indicates a process that has priority above
' IDLE_PRIORITY_CLASS but below NORMAL_PRIORITY_CLASS.
Private Const BELOW_NORMAL_PRIORITY_CLASS As Long = &H4000

' Indicates a process that performs time-critical tasks that
' must be executed immediately for it to run correctly.
' The threads of a high-priority class process preempt the threads
' of normal or idle priority class processes.
Private Const HIGH_PRIORITY_CLASS As Long = &H80

' Indicates a process whose threads run only when the system is
' idle and are preempted by the threads of any process
' running in a higher priority class.
Private Const IDLE_PRIORITY_CLASS As Long = &H40

' Indicates a normal process with no special scheduling needs.
Private Const NORMAL_PRIORITY_CLASS As Long = &H20

' Vista/Server2008: Begin background processing mode. The system lowers
' the resource scheduling priorities of the process (and its threads)
' so that it can perform background work without significantly
' affecting activity in the foreground.
' This value can be specified only if hProcess is a handle
' to the current process. The function fails if the process
' is already in background processing mode.
Private Const PROCESS_MODE_BACKGROUND_BEGIN As Long = &H100000

' Vista/Server2008: End background processing mode. The system restores
' the resource scheduling priorities of the process (and its threads)
' as they were before the process entered background processing mode.
' This value can be specified only if hProcess is a handle
' to the current process. The function fails if the process
' is not in background processing mode.
Private Const PROCESS_MODE_BACKGROUND_END As Long = &H200000

' Indicates a process that has the highest possible priority.
' The threads of a real-time priority class process preempt the
' threads of all other processes, including operating system
' processes performing important tasks.
Private Const REALTIME_PRIORITY_CLASS As Long = &H100

Public Enum TPriorityClassValue
   pcvIdle = IDLE_PRIORITY_CLASS
   pcvBackgroundBegin = PROCESS_MODE_BACKGROUND_BEGIN
   pcvBackgroundEnd = PROCESS_MODE_BACKGROUND_END
   pcvBelowNormal = BELOW_NORMAL_PRIORITY_CLASS
   pcvNormal = NORMAL_PRIORITY_CLASS
   pcvAboveNormal = ABOVE_NORMAL_PRIORITY_CLASS
   pcvHigh = HIGH_PRIORITY_CLASS
   pcvRealtime = REALTIME_PRIORITY_CLASS
End Enum

' The return value is a pseudo handle to the current process.
Private Declare PtrSafe Function GetCurrentProcess Lib "kernel32" () As Long

' Function returns the priority class for the specified process.
Private Declare PtrSafe Function GetPriorityClass Lib "kernel32" ( _
       ByVal hProcess As Long) _
       As Long

' Function sets the priority class for the specified process.
' This value together with the priority value of each thread
' of the process determines each thread's base priority level.
Private Declare PtrSafe Function SetPriorityClass Lib "kernel32" ( _
        ByVal hProcess As Long, _
        ByVal dwPriorityClass As Long) _
        As Long

'
' Private methods
'
Private Sub HandleAPIError(ByVal classErrorNumber As Long, ByRef errorFormatText As String, ByRef actionName As String, ByVal apiErrorNumber As Long)
   Dim mm As New MessageManager

   Err.Raise classErrorNumber, _
             CLASS_NAME, _
             mm.FormatMessageWithParameters(errorFormatText, actionName, apiErrorNumber, mm.GetMessageForWindowsErrorCode(apiErrorNumber))
End Sub

Private Sub CheckPriorityClass(ByVal aPriorityClassValue As TPriorityClassValue)
   Select Case aPriorityClassValue
      Case pcvIdle, pcvBackgroundBegin, pcvBackgroundEnd, pcvBelowNormal, pcvNormal, pcvAboveNormal, pcvHigh, pcvRealtime
         ' Value is valid. So just continue

      Case Else
         Err.Raise ERROR_NUMBER_INVALID_PRIORITY_CLASS, _
                   CLASS_NAME, _
                   ERROR_TEXT_INVALID_PRIORITY_CLASS & Format$(aPriorityClassValue) & " (" & Hex$(aPriorityClassValue) & ")"
   End Select
End Sub

'
' Public methods
'
Public Sub ChangeThisProcessPriority(ByVal newPriorityClassValue As TPriorityClassValue)
   Dim hProcess As Long
   Dim rc As Long
   Dim callError As Long

   CheckPriorityClass newPriorityClassValue

   hProcess = GetCurrentProcess

   ' Get the current Priority
   rc = GetPriorityClass(hProcess)

   ' If the priority specified is not the same as
   ' what's returned by GetPriorityClass
   If rc <> ERROR_FAILURE Then
      If rc <> newPriorityClassValue Then
         ' then attempt to set the new priority for the process
         rc = SetPriorityClass(hProcess, newPriorityClassValue)

         If rc = ERROR_FAILURE Then
            callError = Err.LastDllError
 
            HandleAPIError ERROR_NUMBER_CAN_NOT_SET_PRIORITY_CLASS, _
                           ERROR_MESSAGE_FOR_API, _
                           "set", _
                           callError
         End If
      End If
   Else
      callError = Err.LastDllError
  
      HandleAPIError ERROR_NUMBER_CAN_NOT_SET_PRIORITY_CLASS, _
                     ERROR_MESSAGE_FOR_API, _
                     "get", _
                     callError
   End If
End Sub
