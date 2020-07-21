# VBAUtilities

This is a collection of VBA utilities that I wrote in the last 20 years. I did this because I really love simplicity and elegance and far too many code examples on the internet have neither of these. 

So, I hope you will find these utilities useful. Use them as you wish. Just include a copyright notice that references me. enjoy.

The utilities are categorized in the following way:

## Access

### AccessLockFileReader

A class to read an access lock file in order to find out which computer is holding a lock on the access database.This can either be the database that this class is part of or another database file. It supports mdb and accdb files.

### DAOPropertyManager

Easily set and get DAO properties.

### DBBackupManager

Make a copy of a database and keep a specified number of copied files.

### DBCompressor

Manage compression of an Access database.

### DBTableLinkHelper

Get and change the path to a linked table.

### StatusLine

Show and clear an Access status line. The status line is automatically cleared when the class is destroyed.

## Crypto

### HashCng

A universal hashing class. It calculates [SHA-1](https://en.wikipedia.org/wiki/SHA-1) and [SHA-2](https://en.wikipedia.org/wiki/SHA-2) hashes (with 256, 384 and 512 bits length) and also [HMAC](https://en.wikipedia.org/wiki/HMAC) values with these hashes. It uses  the Windows CNG (Crypto Next Generation) API, so all calculations are done by Windows.

## ErrorHandling

### MessageManager

## ExcelUtilities

### WorkbookCustomPropertyHandler

### WorksheetCustomPropertyHandler

## FileHandling

### DriveHelper

### FileCompressionManager

### RandomFileName

## Internet

### FTPClient

## Math

### BearingHelper

### SphereDistanceCalculator

### Trigonometrics

## NumberConversion

### Base64Converter

### HexConverter

### RomanNumberConverter

## OSUtilities

### SetPriorityClass

### SystemInformation

## Sorting

### Sorter

### Stack

## StringHandling

### StringBuilder

### UTF8Converter

Converts VBA strings from and to [UTF-8](https://en.wikipedia.org/wiki/UTF-8) encoding. Note that the UTF-8 values are byte arrays, not strings. Everybody who returns UTF-8 encodings in strings is doing something seriously wrong.

## Timing

### HighPrecisionTimer

A high precision timer that uses the Windows Performance Counter which has a resolution better than 0.000001 seconds (1µs). 
