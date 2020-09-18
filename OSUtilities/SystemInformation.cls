VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SystemInformation"
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
' Private Konstanten
'
Private Const MAX_COMPUTERNAME_LENGTH_PLUS_1 As Long = 16
Private Const MAX_USERNAME_LENGTH_PLUS_1     As Long = 257
Private Const MAX_LOCALE_TEXT_LENGTH         As Long = 256

'Private Const LOCALE_CUSTOM_DEFAULT     As Long = &HC00
'Private Const LOCALE_CUSTOM_UI_DEFAULT  As Long = &H1400
'Private Const LOCALE_CUSTOM_UNSPECIFIED As Long = &H1000
'Private Const LOCALE_INVARIANT          As Long = &H7F
'Private Const LOCALE_SYSTEM_DEFAULT     As Long = &H800
Private Const LOCALE_USER_DEFAULT       As Long = &H400

Private Const LOCALE_ILANGUAGE = &H1
Private Const LOCALE_SLANGUAGE = &H2
Private Const LOCALE_SENGLANGUAGE = &H1001
Private Const LOCALE_SABBREVLANGNAME = &H3
Private Const LOCALE_SNATIVELANGNAME = &H4
Private Const LOCALE_ICOUNTRY = &H5
Private Const LOCALE_SCOUNTRY = &H6
Private Const LOCALE_SENGCOUNTRY = &H1002
Private Const LOCALE_SABBREVCTRYNAME = &H7
Private Const LOCALE_SNATIVECTRYNAME = &H8
Private Const LOCALE_IDEFAULTLANGUAGE = &H9
Private Const LOCALE_IDEFAULTCOUNTRY = &HA
Private Const LOCALE_IDEFAULTCODEPAGE = &HB
Private Const LOCALE_SLIST = &HC
Private Const LOCALE_IMEASURE = &HD
Private Const LOCALE_SDECIMAL = &HE
Private Const LOCALE_STHOUSAND = &HF
Private Const LOCALE_SGROUPING = &H10
Private Const LOCALE_IDIGITS = &H11
Private Const LOCALE_ILZERO = &H12
Private Const LOCALE_SNATIVEDIGITS = &H13
Private Const LOCALE_SCURRENCY = &H14
Private Const LOCALE_SINTLSYMBOL = &H15
Private Const LOCALE_SMONDECIMALSEP = &H16
Private Const LOCALE_SMONTHOUSANDSEP = &H17
Private Const LOCALE_SMONGROUPING = &H18
Private Const LOCALE_ICURRDIGITS = &H19
Private Const LOCALE_IINTLCURRDIGITS = &H1A
Private Const LOCALE_ICURRENCY = &H1B
Private Const LOCALE_INEGCURR = &H1C
Private Const LOCALE_SDATE = &H1D
Private Const LOCALE_STIME = &H1E
Private Const LOCALE_SSHORTDATE = &H1F
Private Const LOCALE_SLONGDATE = &H20
Private Const LOCALE_STIMEFORMAT = &H1003
Private Const LOCALE_IDATE = &H21
Private Const LOCALE_ILDATE = &H22
Private Const LOCALE_ITIME = &H23
Private Const LOCALE_ICENTURY = &H24
Private Const LOCALE_ITLZERO = &H25
Private Const LOCALE_IDAYLZERO = &H26
Private Const LOCALE_IMONLZERO = &H27
Private Const LOCALE_S1159 = &H28
Private Const LOCALE_S2359 = &H29
Private Const LOCALE_SDAYNAME1 = &H2A
Private Const LOCALE_SDAYNAME2 = &H2B
Private Const LOCALE_SDAYNAME3 = &H2C
Private Const LOCALE_SDAYNAME4 = &H2D
Private Const LOCALE_SDAYNAME5 = &H2E
Private Const LOCALE_SDAYNAME6 = &H2F
Private Const LOCALE_SDAYNAME7 = &H30
Private Const LOCALE_SABBREVDAYNAME1 = &H31
Private Const LOCALE_SABBREVDAYNAME2 = &H32
Private Const LOCALE_SABBREVDAYNAME3 = &H33
Private Const LOCALE_SABBREVDAYNAME4 = &H34
Private Const LOCALE_SABBREVDAYNAME5 = &H35
Private Const LOCALE_SABBREVDAYNAME6 = &H36
Private Const LOCALE_SABBREVDAYNAME7 = &H37
Private Const LOCALE_SMONTHNAME1 = &H38
Private Const LOCALE_SMONTHNAME2 = &H39
Private Const LOCALE_SMONTHNAME3 = &H3A
Private Const LOCALE_SMONTHNAME4 = &H3B
Private Const LOCALE_SMONTHNAME5 = &H3C
Private Const LOCALE_SMONTHNAME6 = &H3D
Private Const LOCALE_SMONTHNAME7 = &H3E
Private Const LOCALE_SMONTHNAME8 = &H3F
Private Const LOCALE_SMONTHNAME9 = &H40
Private Const LOCALE_SMONTHNAME10 = &H41
Private Const LOCALE_SMONTHNAME11 = &H42
Private Const LOCALE_SMONTHNAME12 = &H43
Private Const LOCALE_SABBREVMONTHNAME1 = &H44
Private Const LOCALE_SABBREVMONTHNAME2 = &H45
Private Const LOCALE_SABBREVMONTHNAME3 = &H46
Private Const LOCALE_SABBREVMONTHNAME4 = &H47
Private Const LOCALE_SABBREVMONTHNAME5 = &H48
Private Const LOCALE_SABBREVMONTHNAME6 = &H49
Private Const LOCALE_SABBREVMONTHNAME7 = &H4A
Private Const LOCALE_SABBREVMONTHNAME8 = &H4B
Private Const LOCALE_SABBREVMONTHNAME9 = &H4C
Private Const LOCALE_SABBREVMONTHNAME10 = &H4D
Private Const LOCALE_SABBREVMONTHNAME11 = &H4E
Private Const LOCALE_SABBREVMONTHNAME12 = &H4F
Private Const LOCALE_SABBREVMONTHNAME13 = &H100F
Private Const LOCALE_SPOSITIVESIGN = &H50
Private Const LOCALE_SNEGATIVESIGN = &H51
Private Const LOCALE_IPOSSIGNPOSN = &H52
Private Const LOCALE_INEGSIGNPOSN = &H53
Private Const LOCALE_IPOSSYMPRECEDES = &H54
Private Const LOCALE_IPOSSEPBYSPACE = &H55
Private Const LOCALE_INEGSYMPRECEDES = &H56
Private Const LOCALE_INEGSEPBYSPACE = &H57

Private Const RETURN_ERROR As Long = 0

Private Const UNKOWN As String = "<Unkown>"

'
' References
'
Private Declare PtrSafe Function GetComputerName Lib "kernel32.dll" _
        Alias "GetComputerNameW" ( _
        ByVal lpBuffer As LongPtr, _
        ByRef nSize As Long) As Long

Private Declare PtrSafe Function GetUserName Lib "advapi32.dll" _
        Alias "GetUserNameW" ( _
        ByVal lpBuffer As LongPtr, _
        ByRef nSize As Long) As Long

Private Declare PtrSafe Function GetLocaleInfo Lib "kernel32" _
        Alias "GetLocaleInfoW" ( _
        ByVal Locale As Long, _
        ByVal LCType As Long, _
        ByVal lpLCData As LongPtr, _
        ByVal cchData As Long) As Long

'
' Private methods
'
Private Function GetLocaleText(ByVal localeType As Long) As String
   Dim result As String
   Dim length As Long
   Dim rc As Long

   length = GetLocaleInfo(LOCALE_USER_DEFAULT, localeType, ByVal 0, 0)

   result = String$(length, " ")

   rc = GetLocaleInfo(LOCALE_USER_DEFAULT, localeType, StrPtr(result), length)

   If rc = RETURN_ERROR Then
      GetLocaleText = UNKOWN
   Else
      GetLocaleText = Left$(result, rc - 1)
   End If
End Function

'
' Public methods
'
Public Function GetNameOfComputer() As String
   Dim result As String
   Dim rc As Long
   Dim length As Long

   result = Space$(MAX_COMPUTERNAME_LENGTH_PLUS_1)

   length = Len(result)
   
   rc = GetComputerName(StrPtr(result), length)

   If rc = RETURN_ERROR Then
      GetNameOfComputer = UNKOWN
   Else
      GetNameOfComputer = Left$(result, length)
   End If
End Function

Public Function GetNameOfUser() As String
   Dim result As String
   Dim rc As Long
   Dim length As Long

   result = Space$(MAX_USERNAME_LENGTH_PLUS_1)
   length = Len(result)
   
   rc = GetUserName(StrPtr(result), length)

   If rc = RETURN_ERROR Then
      GetNameOfUser = UNKOWN
   Else
      GetNameOfUser = Left$(result, length - 1)
   End If
End Function

Public Function GetDateSeparator() As String
   GetDateSeparator = GetLocaleText(LOCALE_SDATE)
End Function

Public Function GetDecimalSeparator() As String
   GetDecimalSeparator = GetLocaleText(LOCALE_SDECIMAL)
End Function

Public Function GetListSeparator() As String
   GetListSeparator = GetLocaleText(LOCALE_SLIST)
End Function

Public Function GetTimeSeparator() As String
   GetTimeSeparator = GetLocaleText(LOCALE_STIME)
End Function

Public Function GetMeasure() As String
   If GetLocaleText(LOCALE_IMEASURE) = "0" Then
      GetMeasure = "cm"
   Else
      GetMeasure = "in"
   End If
End Function
