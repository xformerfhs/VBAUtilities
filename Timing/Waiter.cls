VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Waiter"
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
'| Class               | Waiter
'|---------------------+---------------------------------------------------
'| Description         | Suspend program operation for a specified
'|                     | amount of time.
'|---------------------+---------------------------------------------------
'| Author              | Frank Schwab
'|---------------------+---------------------------------------------------
'| Version             | 1.0.0
'|---------------------+---------------------------------------------------
'| Changes             | 2020-08-24  Created. fhs
'+-------------------------------------------------------------------------

Option Explicit

'
' API declarations
'
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'
' Public methods
'

'
'+--------------------------------------------------------------------------
'| Method           | WaitFor
'|------------------+-------------------------------------------------------
'| Purpose          | Suspends program execution for a specfified amount
'|                  | of time
'|------------------+-------------------------------------------------------
'| Parameter        | waitTimeInSeconds: No. of seconds to wait
'|------------------+-------------------------------------------------------
'| Return value     | ./.
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2020-08-24  Created. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | This implementation combines two wait methods:
'|                  |
'|                  |    1. There is a loop for every second that should be waited (VBA wait)
'|                  |    2. 1 second is waited for by calling the sleep API function (OS wait)
'|                  |
'|                  | Why is this?
'|                  |
'|                  | If the wait is done purely in VBA Access has an extremely high CPU usage
'|                  | as it constantly polls for the current time. This is not desirable.
'|                  |
'|                  | If the wait is done purely in the OS the whole Access application hangs while waiting
'|                  | and does not respond to anything.
'|                  |
'|                  | Combining these two wait mechanisms lowers the CPU usage and leaves Access responsive
'|                  | with a maximum delay of 1 second.
'|------------------+-------------------------------------------------------
'| Typical call     | Dim wtr As New Waiter
'|                  | ...
'|                  |    wtr.WaitFor 10   ' Wait 10 seconds
'|                  | ...
'+--------------------------------------------------------------------------
Public Sub WaitFor(ByVal waitTimeInSeconds As Integer)
   Dim secondsAlreadyWaited As Integer

   secondsAlreadyWaited = 0

   Do While secondsAlreadyWaited < waitTimeInSeconds
      Sleep 1000  ' Wait 1 second

      DoEvents ' Process events so Access is responsive

      secondsAlreadyWaited = secondsAlreadyWaited + 1
   Loop
End Sub

'
'+--------------------------------------------------------------------------
'| Method           | WaitForRandomTimeInInterval
'|------------------+-------------------------------------------------------
'| Purpose          | Suspends program execution for a random amount
'|                  | of time
'|------------------+-------------------------------------------------------
'| Parameters       | minWaitTime: Minimum wait time in seconds
'|                  | maxWaitTime: Minimum wait time in seconds
'|------------------+-------------------------------------------------------
'| Return value     | ./.
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2020-08-24  Created. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | ./.
'|------------------+-------------------------------------------------------
'| Typical call     | Dim wtr As New Waiter
'|                  | ...
'|                  |    wtr.WaitForRandomTimeInInterval 10, 20   ' Wait between 10 and 20 seconds
'|                  | ...
'+--------------------------------------------------------------------------
Public Sub WaitForRandomTimeInInterval(ByVal minWaitTime As Integer, ByVal maxWaitTime As Integer)
   Me.WaitFor CInt(Fix((Rnd() * (maxWaitTime - minWaitTime + 1)))) + minWaitTime
End Sub

'
' Class methods
'
Private Sub Class_Initialize()
   Randomize
End Sub
