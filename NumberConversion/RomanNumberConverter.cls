VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RomanNumberConverter"
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
'| Class               | RomanNumberConverter
'|---------------------+---------------------------------------------------
'| Description         | Converts integers from and to roman number representation
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
' This class is based on the functions found here:
' http://www.vb-helper.com/howto_roman_arabic.html
'

'
' Private constants
'
Private Const VALID_ARABIC_NUMERALS As String = "-0123456789"

'
' Constants for exceptions
'
Private Const MODULE_NAME As String = "RomanNumberConverter"
Private Const ERR_NUMBER_START As Long = vbObjectError + 32300

Private Const ERR_INVALID_ROMAN_NUMERAL As Long = ERR_NUMBER_START
Private Const STR_ERR_INVALID_ROMAN_NUMERAL As String = "Invalid roman numeral: "

Private Const ERR_INVALID_ARABIC_NUMERAL As Long = ERR_NUMBER_START + 1
Private Const STR_ERR_INVALID_ARABIC_NUMERAL As String = "Invalid arabic numeral: "

'
' Private methods
'
Private Function IsValidInteger(ByRef aString As String) As Boolean
   Dim i As Long
   Dim pos As Long

   IsValidInteger = True

   For i = 1 To Len(aString)
      pos = InStr(1, VALID_ARABIC_NUMERALS, Mid$(aString, i, 1))

      If pos = 0 Then
         IsValidInteger = False

         Exit Function
      Else
         If pos = 1 Then
            If i > 1 Then
               IsValidInteger = False

               Exit Function
            End If
         End If
      End If
   Next i
End Function

' Add appropriate Roman digits to the result.
' The tenLetter, fiveLetter, and oneLetter
' are the digits for 10, 5, and 1 at this
' power of ten. For example, 10/5/1 = X/V/I,
' 100/50/10 = C/L/X, etc.
Private Function GetLowRomanDigits(ByVal numberToConvert As Long, _
                                   ByVal tenLetter As String, _
                                   ByVal fiveLetter As String, _
                                   ByVal oneLetter As String) As String
   Dim result As String

   Select Case numberToConvert
      Case 1 To 3
          result = String$(numberToConvert, oneLetter)

      Case 4
          result = oneLetter & fiveLetter

      Case 5
          result = fiveLetter

      Case 6 To 8
          result = fiveLetter & String$(numberToConvert - 5, oneLetter)

      Case 9
          result = oneLetter & tenLetter
   End Select

   GetLowRomanDigits = result
End Function

'
' Public methods
'
Public Function RomanToArabicNumber(ByVal romanNumber As String) As Long
   Dim i As Long

   Dim ch As String * 1

   Dim result As Long
   Dim new_value As Long
   Dim old_value As Long

   old_value = 1000

   For i = 1 To Len(romanNumber)
      ' See what the next character is worth.
      ch = UCase$(Mid$(romanNumber, i, 1))

      Select Case ch
         Case "I"
            new_value = 1

         Case "V"
            new_value = 5

         Case "X"
            new_value = 10

         Case "L"
            new_value = 50

         Case "C"
            new_value = 100

         Case "D"
            new_value = 500

         Case "M"
             new_value = 1000

         Case Else
            Err.Raise ERR_INVALID_ROMAN_NUMERAL, _
                      MODULE_NAME, _
                      STR_ERR_INVALID_ROMAN_NUMERAL & ch
      End Select

      result = result + new_value

      ' See if this character is bigger
      ' than the previous one.
      If new_value > old_value Then
          ' The new value > the previous one.
          ' Add this value to the result
          ' and subtract the previous one twice.
          result = result - old_value - old_value
      End If

      old_value = new_value
   Next i

   RomanToArabicNumber = result
End Function

Public Function RomanToArabicString(ByRef roman_string As String) As String
   RomanToArabicString = Format$(Me.RomanToArabicNumber(roman_string))
End Function

Public Function ArabicNumberToRoman(ByVal arabicNumber As Long) As String
   Dim quotient As Long
   Dim result As String

   If arabicNumber > 0 Then
      ' Pull out thousands.
      quotient = arabicNumber \ 1000
      arabicNumber = arabicNumber - quotient * 1000
      result = String$(quotient, "M")

      ' Pull out hundreds.
      quotient = arabicNumber \ 100
      arabicNumber = arabicNumber - quotient * 100
      result = result & GetLowRomanDigits(quotient, "M", "D", "C")
   
      ' Pull out tens.
      quotient = arabicNumber \ 10
      arabicNumber = arabicNumber - quotient * 10
      result = result & GetLowRomanDigits(quotient, "C", "L", "X")
   
      ' Pull out ones.
      result = result & GetLowRomanDigits(arabicNumber, "X", "V", "I")
   Else
      If arabicNumber = 0 Then
         result = "0"
      Else
         result = "-" & Me.ArabicNumberToRoman(-arabicNumber)
      End If
   End If

   ArabicNumberToRoman = result
End Function

Public Function ArabicStringToRoman(ByVal arabicString As String) As String
   If Len(arabicString) > 0 Then
      If IsValidInteger(arabicString) Then
         ArabicStringToRoman = Me.ArabicNumberToRoman(CLng(arabicString))
      Else
         Err.Raise ERR_INVALID_ARABIC_NUMERAL, _
                   MODULE_NAME, _
                   STR_ERR_INVALID_ARABIC_NUMERAL & arabicString
      End If
   Else
      ArabicStringToRoman = ""
   End If
End Function
