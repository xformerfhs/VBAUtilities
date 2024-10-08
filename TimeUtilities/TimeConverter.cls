VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TimeConverter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
'+-------------------------------------------------------------------------
'|
'| SPDX-FileCopyrightText: 2022-2024 Frank Schwab
'|
'| SPDX-License-Identifier: MIT
'|
'| Copyright 2022-2024, Frank Schwab
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
'| Class               | TimeConverter
'|---------------------+---------------------------------------------------
'| Description         | Convert VBA timestamps from and to Unix time
'|                     | and local time between UTC time and vice versa.
'|---------------------+---------------------------------------------------
'| Author              | Frank Schwab
'|---------------------+---------------------------------------------------
'| Version             | 2.0.0
'|---------------------+---------------------------------------------------
'| Changes             | 2020-08-24  Created. fhs
'|                     | 2022-01-24  Completely rewritten. fhs
'+-------------------------------------------------------------------------

Option Explicit

'
' Constants for error messages
'

Private Const STR_ERROR_SOURCE   As String = "TimeConverter"
Private Const ERR_BASE As Long = vbObjectError + 22617

' Crypto API error
Private Const STR_ERR_WINDOWS_API As String = "API function %1 returned code 0x%2: %3"
Private Const NUM_ERR_WINDOWS_API As Long = ERR_BASE

Private Const STR_API_TIME_ZONE_INFORMATION As String = "GetTimeZoneInformation"

'
' Private constants
'

Private Const UNIX_TIME_START As Double = 25569# ' DateSerial(1970, 1, 1)

Private Const MINUTES_PER_DAY As Double = 24# * 60#
Private Const SECONDS_PER_DAY As Double = MINUTES_PER_DAY * 60#

Private Const MINUTES_PER_HOUR As Long = 60&

Private Const RC_ERROR As Long = 0

'
' System data types
'
Private Type SYSTEMTIME
        wYear As Integer
        wMonth As Integer
        wDayOfWeek As Integer
        wDay As Integer
        wHour As Integer
        wMinute As Integer
        wSecond As Integer
        wMilliseconds As Integer
End Type

Private Type TIME_ZONE_INFORMATION
        Bias As Long
        StandardName(0 To 31) As Integer
        StandardDate As SYSTEMTIME
        StandardBias As Long
        DaylightName(0 To 31) As Integer
        DaylightDate As SYSTEMTIME
        DaylightBias As Long
End Type

Private Enum TIME_ZONE_ID
   TIME_ZONE_ID_INVALID = -1
   TIME_ZONE_ID_UNKNOWN = 0
   TIME_ZONE_ID_STANDARD = 1
   TIME_ZONE_ID_DAYLIGHT = 2
End Enum

'
' System API declarations
'
Private Declare PtrSafe Function GetTimeZoneInformation Lib "kernel32" _
       (ByRef lpTimeZoneInformation As TIME_ZONE_INFORMATION) As Long

Private Declare PtrSafe Function SystemTimeToTzSpecificLocalTime Lib "kernel32" _
       (ByRef lpTimeZoneInformation As TIME_ZONE_INFORMATION, _
        ByRef lpUniversalTime As SYSTEMTIME, _
        ByRef lpLocalTime As SYSTEMTIME) As Long

Private Declare PtrSafe Function TzSpecificLocalTimeToSystemTime Lib "kernel32" _
       (ByRef lpTimeZoneInformation As TIME_ZONE_INFORMATION, _
        ByRef lpLocalTime As SYSTEMTIME, _
        ByRef lpUniversalTime As SYSTEMTIME) As Long

'
' Private methods
'

'
'+--------------------------------------------------------------------------
'| Method           | HandleError
'|------------------+-------------------------------------------------------
'| Description      | Handle Windows API errors by throwing an exception
'|------------------+-------------------------------------------------------
'| Parameter        | apiFunctionName: Name of the failing API function
'|                  | rc: Status code of the failing API function
'|------------------+-------------------------------------------------------
'| Return values    | ./.
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2022-01-24  Created. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | This methods always throws an exception.
'+--------------------------------------------------------------------------
'
Private Sub HandleError(ByRef apiFunctionName As String, _
                        ByVal rc As Long)
   Dim mm As New MessageManager
   
   Err.Raise NUM_ERR_WINDOWS_API, _
             STR_ERROR_SOURCE, _
             mm.FormatMessageWithParameters(STR_ERR_WINDOWS_API, _
                                            apiFunctionName, _
                                            Hex$(rc), _
                                            mm.GetMessageForNTStatusCode(rc))
End Sub

'
'+--------------------------------------------------------------------------
'| Method           | ConvertSystemTimeToDate
'|------------------+-------------------------------------------------------
'| Description      | Convert a SYSTEMTIME to VBA Date.
'|------------------+-------------------------------------------------------
'| Parameter        | forSystemTime: The SYSTEMTIME to convert.
'|------------------+-------------------------------------------------------
'| Return values    | VBA Date value of the supplied SYSTEMTIME.
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2020-08-24  Created. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | ./.
'+--------------------------------------------------------------------------
'
Private Function ConvertSystemTimeToDate(ByRef forSystemTime As SYSTEMTIME) As Date
   With forSystemTime
      ConvertSystemTimeToDate = DateSerial(.wYear, .wMonth, .wDay) + _
                              TimeSerial(.wHour, .wMinute, .wSecond)
   End With
End Function

'
'+--------------------------------------------------------------------------
'| Method           | ConvertDateToSystemTime
'|------------------+-------------------------------------------------------
'| Description      | Convert a VBA Date to a SYSTEMTIME.
'|------------------+-------------------------------------------------------
'| Parameter        | forDate: The Date to convert.
'|------------------+-------------------------------------------------------
'| Return values    | SYSTEMTIME value of the supplied VBA Date.
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2020-08-24  Created. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | ./.
'+--------------------------------------------------------------------------
'
Private Function ConvertDateToSystemTime(ByVal forDate As Date) As SYSTEMTIME
   Dim result As SYSTEMTIME

   With result
      .wDay = Day(forDate)
      .wMonth = Month(forDate)
      .wYear = Year(forDate)
      .wHour = Hour(forDate)
      .wMinute = Minute(forDate)
      .wSecond = Second(forDate)
      .wMilliseconds = 0
      .wDayOfWeek = Weekday(forDate, vbSunday) - 1
   End With

   ConvertDateToSystemTime = result
End Function


'
' Public methods
'

'
'+--------------------------------------------------------------------------
'| Method           | ConvertUTCToLocalTime
'|------------------+-------------------------------------------------------
'| Description      | Convert an UTC to local time.
'|------------------+-------------------------------------------------------
'| Parameter        | utcTime: UTC time to convert.
'|------------------+-------------------------------------------------------
'| Return values    | Local time for the supplied UTC.
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2020-08-24  Created. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | ./.
'+--------------------------------------------------------------------------
'
Public Function ConvertUTCToLocalTime(ByVal utcTime As Date) As Date
   Dim TZI As TIME_ZONE_INFORMATION
   Dim DST As TIME_ZONE_ID

   Dim utcSystemTime As SYSTEMTIME
   Dim localSystemTime As SYSTEMTIME
   
   Dim rc As Long
   
   utcSystemTime = ConvertDateToSystemTime(utcTime)

   DST = GetTimeZoneInformation(TZI)

   If DST = TIME_ZONE_ID_INVALID Then
      HandleError STR_API_TIME_ZONE_INFORMATION, Err.LastDllError
   Else
      If DST <> TIME_ZONE_ID_UNKNOWN Then
         rc = SystemTimeToTzSpecificLocalTime(TZI, utcSystemTime, localSystemTime)

         If rc <> RC_ERROR Then
            ConvertUTCToLocalTime = ConvertSystemTimeToDate(localSystemTime)
         Else
            HandleError "SystemTimeToTzSpecificLocalTime", Err.LastDllError
         End If
      Else
         ConvertUTCToLocalTime = utcTime
      End If
   End If
End Function

'
'+--------------------------------------------------------------------------
'| Method           | ConvertLocalTimeToUTC
'|------------------+-------------------------------------------------------
'| Description      | Convert a local time to UTC.
'|------------------+-------------------------------------------------------
'| Parameter        | localTime: Local time to convert.
'|------------------+-------------------------------------------------------
'| Return values    | UTC for the supplied local time.
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2020-08-24  Created. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | ./.
'+--------------------------------------------------------------------------
'
Public Function ConvertLocalTimeToUTC(ByVal localTime As Date) As Date
   Dim TZI As TIME_ZONE_INFORMATION
   Dim DST As TIME_ZONE_ID

   Dim utcSystemTime As SYSTEMTIME
   Dim localSystemTime As SYSTEMTIME
   
   Dim rc As Long
   
   localSystemTime = ConvertDateToSystemTime(localTime)

   DST = GetTimeZoneInformation(TZI)

   If DST = TIME_ZONE_ID_INVALID Then
      HandleError STR_API_TIME_ZONE_INFORMATION, Err.LastDllError
   Else
      If DST <> TIME_ZONE_ID_UNKNOWN Then
         rc = TzSpecificLocalTimeToSystemTime(TZI, localSystemTime, utcSystemTime)

         If rc <> RC_ERROR Then
            ConvertLocalTimeToUTC = ConvertSystemTimeToDate(utcSystemTime)
         Else
            HandleError "TzSpecificLocalTimeToSystemTime", Err.LastDllError
         End If
      Else
         ConvertLocalTimeToUTC = localTime
      End If
   End If
End Function

'
'+--------------------------------------------------------------------------
'| Method           | ConvertDateToUnixTime
'|------------------+-------------------------------------------------------
'| Description      | Convert a VBA Date to Unix time.
'|------------------+-------------------------------------------------------
'| Parameter        | vbaTime: VBA Date to convert.
'|------------------+-------------------------------------------------------
'| Return values    | UNIX timestamp for the supplied VBA Date.
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2020-08-24  Created. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | ./.
'+--------------------------------------------------------------------------
'
Public Function ConvertDateToUnixTime(ByVal vbaTime As Date) As Long
   ConvertDateToUnixTime = (vbaTime - UNIX_TIME_START) * SECONDS_PER_DAY
End Function

'
'+--------------------------------------------------------------------------
'| Method           | ConvertUnixTimeToDate
'|------------------+-------------------------------------------------------
'| Description      | Convert an Unix time to VBA Date.
'|------------------+-------------------------------------------------------
'| Parameter        | unixTime: UNIX timestamp to convert.
'|------------------+-------------------------------------------------------
'| Return values    | VBA Date for the supplied UNIX timestamp.
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2020-08-24  Created. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | ./.
'+--------------------------------------------------------------------------
'
Public Function ConvertUnixTimeToDate(ByVal unixTime As Long) As Date
   ConvertUnixTimeToDate = (unixTime / SECONDS_PER_DAY) + UNIX_TIME_START
End Function

'
'+--------------------------------------------------------------------------
'| Method           | GetLocalTimeOffsetMinutes
'|------------------+-------------------------------------------------------
'| Description      | Gets the offset of a local time from UTC in minutes.
'|------------------+-------------------------------------------------------
'| Parameter        | localTime: Local time to check.
'|------------------+-------------------------------------------------------
'| Return values    | Offset of the supplied local time from UTC in minutes.
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2022-01-24  Created. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | This method calculates LocalTime - UTCTime.
'+--------------------------------------------------------------------------
'
Public Function GetLocalTimeOffsetMinutes(ByVal localTime As Date) As Long
   Dim utcTime As Date

   utcTime = ConvertLocalTimeToUTC(localTime)

   GetLocalTimeOffsetMinutes = CLng(Round((localTime - utcTime) * MINUTES_PER_DAY))
End Function

'
'+--------------------------------------------------------------------------
'| Method           | MinutesAsOffset
'|------------------+-------------------------------------------------------
'| Description      | Converts offset minutes to an offset expression.
'|------------------+-------------------------------------------------------
'| Parameter        | minutes: Offset in minutes.
'|------------------+-------------------------------------------------------
'| Return values    | A string of the form "+HH:MM" for the supplied minutes.
'|------------------+-------------------------------------------------------
'| Author           | Frank Schwab
'|------------------+-------------------------------------------------------
'| Changes          | 2022-01-24  Created. fhs
'|------------------+-------------------------------------------------------
'| Remarks          | ./.
'+--------------------------------------------------------------------------
'
Public Function MinutesAsOffset(ByVal minutes As Long) As String
   Dim prefix As String * 1

   If minutes >= 0 Then
      prefix = "+"
   Else
      prefix = "-"
   End If

   Dim absMinutes As Long

   absMinutes = Abs(minutes)

   Dim hours As Long

   hours = absMinutes \ MINUTES_PER_HOUR

   MinutesAsOffset = prefix & Format$(hours, "00") & ":" & Format$(minutes - (hours * MINUTES_PER_HOUR), "00")
End Function
