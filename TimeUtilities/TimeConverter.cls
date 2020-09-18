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
'| Class               | TimeConverter
'|---------------------+---------------------------------------------------
'| Description         | Convert VBA timestamps from and to Unix time
'|                     | and local time between UTC time and vice versa.
'|---------------------+---------------------------------------------------
'| Author              | Frank Schwab
'|---------------------+---------------------------------------------------
'| Version             | 1.0.0
'|---------------------+---------------------------------------------------
'| Changes             | 2020-08-24  Created. fhs
'+-------------------------------------------------------------------------

Option Explicit

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
   TIME_ZONE_ID_UNKNOWN = 0
   TIME_ZONE_ID_STANDARD = 1
   TIME_ZONE_ID_DAYLIGHT = 2
End Enum

'
' System API declarations
'
Private Declare PtrSafe Function GetTimeZoneInformation Lib "kernel32" _
       (lpTimeZoneInformation As TIME_ZONE_INFORMATION) As Long

Private Declare PtrSafe Function SystemTimeToTzSpecificLocalTime Lib "kernel32" _
       (ByRef lpTimeZoneInformation As TIME_ZONE_INFORMATION, _
        ByRef lpUniversalTime As SYSTEMTIME, _
        ByRef lpLocalTime As SYSTEMTIME) As Long

Private Declare PtrSafe Function TzSpecificLocalTimeToSystemTime Lib "kernel32" _
       (ByRef lpTimeZoneInformation As TIME_ZONE_INFORMATION, _
        ByRef lpLocalTime As SYSTEMTIME, _
        ByRef lpUniversalTime As SYSTEMTIME) As Long

Private Const UNIX_TIME_START As Double = 25569# ' DateSerial(1970, 1, 1)

Private Const SECONDS_PER_DAY As Double = 24# * 60# * 60#

'
' Private methods
'
Private Function GetDateFromSystemTime(ByRef forSystemTime As SYSTEMTIME) As Date
   With forSystemTime
      GetDateFromSystemTime = DateSerial(.wYear, .wMonth, .wDay) + _
                              TimeSerial(.wHour, .wMinute, .wSecond)
   End With
End Function

Private Function GetSystemTimeFromDate(ByVal forDate As Date) As SYSTEMTIME
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

   GetSystemTimeFromDate = result
End Function

'
' Public methods
'
Public Function GetLocalTimeFromUTC(ByVal utcTime As Date) As Date
   Dim TZI As TIME_ZONE_INFORMATION
   Dim DST As TIME_ZONE_ID

   Dim utcSystemTime As SYSTEMTIME
   Dim localSystemTime As SYSTEMTIME
   
   Dim rc As Long
   
   utcSystemTime = GetSystemTimeFromDate(utcTime)

   DST = GetTimeZoneInformation(TZI)
   rc = SystemTimeToTzSpecificLocalTime(TZI, utcSystemTime, localSystemTime)

   GetLocalTimeFromUTC = GetDateFromSystemTime(localSystemTime)
End Function

Public Function GetUTCFromLocalTime(ByVal localTime As Date) As Date
   Dim TZI As TIME_ZONE_INFORMATION
   Dim DST As TIME_ZONE_ID

   Dim utcSystemTime As SYSTEMTIME
   Dim localSystemTime As SYSTEMTIME
   
   Dim rc As Long
   
   localSystemTime = GetSystemTimeFromDate(localTime)

   DST = GetTimeZoneInformation(TZI)
   rc = TzSpecificLocalTimeToSystemTime(TZI, localSystemTime, utcSystemTime)

   GetUTCFromLocalTime = GetDateFromSystemTime(utcSystemTime)
End Function

Public Function GetUnixTimeFromVBATime(ByVal vbaTime As Date) As Long
   GetUnixTimeFromVBATime = (vbaTime - UNIX_TIME_START) * SECONDS_PER_DAY
End Function

Public Function GetVBATimeFromUnixTime(ByVal unixTime As Long) As Date
   GetVBATimeFromUnixTime = (unixTime / SECONDS_PER_DAY) + UNIX_TIME_START
End Function
